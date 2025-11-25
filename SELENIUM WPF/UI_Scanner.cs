using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using Forms = System.Windows.Forms;

namespace SELENIUM_WPF
{
    [Flags]
    public enum PatternFlags
    {
        None = 0,
        Invoke = 1 << 0,
        Value = 1 << 1,
        SelectionItem = 1 << 2,
        ExpandCollapse = 1 << 3,
        Toggle = 1 << 4,
        Scroll = 1 << 5,
        ScrollItem = 1 << 6,
        Text = 1 << 7,
        LegacyIAccessible = 1 << 8
    }

    public sealed class ElementRecord
    {
        // Ссылка на UIA-элемент (для последующих действий)
        public AutomationElement Element { get; init; }

        // Снимок базовых свойств
        public int[] RuntimeId { get; init; }
        public string Name { get; init; }
        public string AutomationId { get; init; }
        public string ClassName { get; init; }
        public string ControlType { get; init; } // "Button", "Edit", ...
        public string FrameworkId { get; init; } // "WPF", "WinForm", ...
        public int ProcessId { get; init; }
        public Rect Bounds { get; init; }
        public bool IsOffscreen { get; init; }
        public bool IsEnabled { get; init; }
        public bool IsControlElement { get; init; }
        public bool IsContentElement { get; init; }

        // Координаты центра для наведения мыши
        public System.Drawing.Point CenterPoint { get; init; }

        // Поддерживаемые паттерны
        public PatternFlags Patterns { get; init; }

        // Результаты оценки доступности
        public bool IsOnAnyScreen { get; set; }
        public bool HasClickablePoint { get; set; }
        public bool AncestorCollapsed { get; set; }
        public bool WindowMinimized { get; set; }
        public bool InteractableNow { get; set; }
        public string NotInteractableReason { get; set; }

        public override string ToString()
        {
            return $"{ControlType,-12} Name='{Name}' Aid='{AutomationId}' Class='{ClassName}' " +
                   $"Bounds=[{(int)Bounds.X},{(int)Bounds.Y} {(int)Bounds.Width}x{(int)Bounds.Height}] " +
                   $"Center=({CenterPoint.X},{CenterPoint.Y}) " +
                   $"Enabled={IsEnabled} Offscreen={IsOffscreen} Patterns={Patterns} " +
                   $"InteractableNow={InteractableNow} {(InteractableNow ? "" : $"Reason={NotInteractableReason}")}";
        }
    }

    public static class UI_Scanner
    {


        //  Снимок всех контролов под корнем, с кэшированием свойств для производительности
        public static List<ElementRecord> SnapshotControls(AutomationElement root)
        {

            if (root == null) throw new ArgumentNullException(nameof(root));

            var cache = new CacheRequest
            {
                TreeScope = TreeScope.Subtree,
                AutomationElementMode = AutomationElementMode.Full // чтобы Element оставался «живым»
            };

            cache.Add(AutomationElement.NameProperty);
            cache.Add(AutomationElement.AutomationIdProperty);
            cache.Add(AutomationElement.ClassNameProperty);
            cache.Add(AutomationElement.ControlTypeProperty);
            cache.Add(AutomationElement.FrameworkIdProperty);
            cache.Add(AutomationElement.ProcessIdProperty);
            cache.Add(AutomationElement.BoundingRectangleProperty);
            cache.Add(AutomationElement.IsOffscreenProperty);
            cache.Add(AutomationElement.IsEnabledProperty);
            cache.Add(AutomationElement.IsControlElementProperty);
            cache.Add(AutomationElement.IsContentElementProperty);
            cache.Add(AutomationElement.RuntimeIdProperty);

            using (cache.Activate()) 
            {
                
                var isControlCondition = new PropertyCondition(AutomationElement.IsControlElementProperty, true);
                var all = root.FindAll(TreeScope.Subtree, isControlCondition);

                var list = new List<ElementRecord>(all.Count);
                foreach (AutomationElement el in all)
                {
                    
                    var ctName = el.Cached.ControlType.ProgrammaticName;
                    var controlTypeShort = ctName != null && ctName.Contains(".")
                        ? ctName[(ctName.IndexOf('.') + 1)..]
                        : ctName;

                    var bounds = el.Cached.BoundingRectangle;
                    var centerPoint = CalculateCenterPoint(bounds);

                    list.Add(new ElementRecord
                    {
                        Element = el,
                        RuntimeId = SafeGetRuntimeId(el),
                        Name = el.Cached.Name,
                        AutomationId = el.Cached.AutomationId,
                        ClassName = el.Cached.ClassName,
                        ControlType = controlTypeShort,
                        FrameworkId = el.Cached.FrameworkId,
                        ProcessId = el.Cached.ProcessId,
                        Bounds = bounds,
                        CenterPoint = centerPoint,
                        IsOffscreen = el.Cached.IsOffscreen,
                        IsEnabled = el.Cached.IsEnabled,
                        IsControlElement = el.Cached.IsControlElement,
                        IsContentElement = el.Cached.IsContentElement,
                        Patterns = DetectPatterns(el)
                    });
                }

                return list;
            }
        }


        //  Оценка «доступности сейчас» для каждого элемента из снимка
        public static void EvaluateAvailability(AutomationElement root, IList<ElementRecord> records)
        {
            if (root == null) throw new ArgumentNullException(nameof(root));
            if (records == null) throw new ArgumentNullException(nameof(records));

            bool windowMinimized = IsWindowMinimized(root);

            foreach (var r in records)
            {
                r.WindowMinimized = windowMinimized;

                // Базовые быстрые отказы
                if (windowMinimized)
                {
                    r.InteractableNow = false;
                    r.NotInteractableReason = "Окно свернуто";
                    continue;
                }

                if (!r.IsEnabled)
                {
                    r.InteractableNow = false;
                    r.NotInteractableReason = "Элемент отключён (IsEnabled=false)";
                    continue;
                }

                if (r.Bounds.IsEmpty || r.Bounds.Width <= 0 || r.Bounds.Height <= 0)
                {
                    r.InteractableNow = false;
                    r.NotInteractableReason = "Нулевые границы";
                    continue;
                }

                // Видимость: не offscreen и попадает на какой-либо экран
                r.IsOnAnyScreen = IntersectsAnyScreen(r.Bounds);
                if (r.IsOffscreen || !r.IsOnAnyScreen)
                {
                    r.InteractableNow = false;
                    r.NotInteractableReason = r.IsOffscreen ? "Вне экрана (IsOffscreen=true)" : "Вне области экранов";
                    continue;
                }

                // Предки не спрятали элемент
                r.AncestorCollapsed = HasCollapsedAncestor(r.Element);
                if (r.AncestorCollapsed)
                {
                    r.InteractableNow = false;
                    r.NotInteractableReason = "Скрыт родителем (Collapse/Collapsed)";
                    continue;
                }

                // Есть ли семантическое действие или хотя бы кликабельная точка
                r.HasClickablePoint = TryHasClickablePoint(r.Element);

                bool hasSemanticAction =
                       r.Patterns.HasFlag(PatternFlags.Invoke)
                    || r.Patterns.HasFlag(PatternFlags.Value)
                    || r.Patterns.HasFlag(PatternFlags.SelectionItem)
                    || r.Patterns.HasFlag(PatternFlags.Toggle)
                    || r.Patterns.HasFlag(PatternFlags.ExpandCollapse);

                if (hasSemanticAction || r.HasClickablePoint)
                {
                    r.InteractableNow = true;
                    r.NotInteractableReason = null;
                }
                else
                {
                    r.InteractableNow = false;
                    r.NotInteractableReason = "Нет семантического паттерна и кликабельной точки";
                }
            }
        }




        //  ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ

        private static System.Drawing.Point CalculateCenterPoint(Rect bounds)
        {
            //  вычисление центра и привидение к целочисленным координатам
            int centerX = (int)Math.Round(bounds.X + bounds.Width / 2);
            int centerY = (int)Math.Round(bounds.Y + bounds.Height / 2);

            return new System.Drawing.Point(centerX, centerY);
        }


        private static int[] SafeGetRuntimeId(AutomationElement el)
        {
            try { return el.GetRuntimeId(); }
            catch { return Array.Empty<int>(); }
        }


        private static PatternFlags DetectPatterns(AutomationElement el)
        {
            PatternFlags flags = PatternFlags.None;

            void Check(AutomationPattern p, PatternFlags f)
            {
                try
                {
                    if (el.TryGetCurrentPattern(p, out _)) flags |= f;
                }
                catch
                {
                    
                }
            }

            Check(InvokePattern.Pattern, PatternFlags.Invoke);
            Check(ValuePattern.Pattern, PatternFlags.Value);
            Check(SelectionItemPattern.Pattern, PatternFlags.SelectionItem);
            Check(ExpandCollapsePattern.Pattern, PatternFlags.ExpandCollapse);
            Check(TogglePattern.Pattern, PatternFlags.Toggle);
            Check(ScrollPattern.Pattern, PatternFlags.Scroll);
            Check(ScrollItemPattern.Pattern, PatternFlags.ScrollItem);
            Check(TextPattern.Pattern, PatternFlags.Text);

            
            try
            {
               
                var legacyType = typeof(AutomationElement).Assembly.GetTypes()
                    .FirstOrDefault(t => t.Name.Contains("LegacyIAccessible") && t.Name.Contains("Pattern"));

                if (legacyType != null)
                {
                    var patternField = legacyType.GetField("Pattern",
                        System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);

                    if (patternField?.GetValue(null) is AutomationPattern legacyPattern)
                    {
                        Check(legacyPattern, PatternFlags.LegacyIAccessible);
                    }
                }
            }
            catch
            {
                
            }

            return flags;
        }


        private static bool IntersectsAnyScreen(Rect r)
        {
            // приведение для сравнения с экраном
            var rr = new System.Drawing.Rectangle((int)r.X, (int)r.Y, (int)Math.Max(0, r.Width), (int)Math.Max(0, r.Height));
            foreach (var screen in Forms.Screen.AllScreens)
            {
                if (rr.IntersectsWith(screen.Bounds)) return true;
            }
            return false;
        }


        private static bool HasCollapsedAncestor(AutomationElement el)
        {
            var walker = TreeWalker.ControlViewWalker;
            var parent = walker.GetParent(el);
            while (parent != null)
            {
                if (parent.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out var p))
                {
                    var state = ((ExpandCollapsePattern)p).Current.ExpandCollapseState;
                    if (state == ExpandCollapseState.Collapsed)
                        return true;
                }
                // Часто скрытые контейнеры помечаются IsOffscreen=true
                if (parent.Current.IsOffscreen) return true;

                parent = walker.GetParent(parent);
            }
            return false;
        }


        private static bool TryHasClickablePoint(AutomationElement el)
        {
            try
            {
                return el.TryGetClickablePoint(out _);
            }
            catch
            {
                
                return false;
            }
        }


        private static bool IsWindowMinimized(AutomationElement root)
        {
            try
            {
                if (root.TryGetCurrentPattern(WindowPattern.Pattern, out var p))
                {
                    var state = ((WindowPattern)p).Current.WindowVisualState;
                    return state == WindowVisualState.Minimized;
                }
            }
            catch {}
            return false;
        }
    }
}