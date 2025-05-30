using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Collections.Generic;
static class Program
{
    [STAThread]
    static void Main()
    {
        System.Windows.Forms.Application.Run(new Inspector());
    }
}


class Inspector : Form
{
    const int WM_HOTKEY = 0x0312;
    const int HOTKEY_ID = 0xABCD;
    const uint MOD_CONTROL = 0x2;
    const uint VK_F = 0x46;

    [DllImport("user32.dll")]
    static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    static extern bool GetCursorPos(out POINT lpPoint);


    private SplitContainer splitContainer;
    private TextBox searchBar;
    private TreeView treeViewAll;
    private TextBox textBoxDetails;
    private TextBox textBoxSuggestions;
    private static readonly Dictionary<string, string> PythonControlMap = new()
    {
        { "app bar",      "AppBarControl"     },
        { "button",       "ButtonControl"     },
        { "calendar",     "CalendarControl"   },
        { "check box",    "CheckBoxControl"   },
        { "combo box",    "ComboBoxControl"   },
        { "custom",       "CustomControl"     },
        { "data grid",    "DataGridControl"   },
        { "data item",    "DataItemControl"   },
        { "document",     "DocumentControl"   },
        { "edit",         "EditControl"       },
        { "group",        "GroupControl"      },
        { "header",       "HeaderControl"     },
        { "header item",  "HeaderItemControl" },
        { "hyperlink",    "HyperlinkControl"  },
        { "image",        "ImageControl"      },
        { "list",         "ListControl"       },
        { "list item",    "ListItemControl"   },
        { "menu bar",     "MenuBarControl"    },
        { "menu",         "MenuControl"       },
        { "menu item",    "MenuItemControl"   },
        { "pane",         "PaneControl"       },
        { "progress bar", "ProgressBarControl"},
        { "radio button", "RadioButtonControl"},
        { "scroll bar",   "ScrollBarControl"  },
        { "slider",       "SliderControl"     },
        { "spinner",      "SpinnerControl"    },
        { "split button", "SplitButtonControl"},
        { "status bar",   "StatusBarControl"  },
        { "tab",          "TabControl"        },
        { "tab item",     "TabItemControl"    },
        { "table",        "TableControl"      },
        { "text",         "TextControl"       },
        { "thumb",        "ThumbControl"      },
        { "title bar",    "TitleBarControl"   },
        { "tool bar",     "ToolBarControl"    },
        { "tool tip",     "ToolTipControl"    },
        { "tree",         "TreeControl"       },
        { "tree item",    "TreeItemControl"   },
        { "window",       "WindowControl"     },
        { "item",         "DataItemControl"   }
    };

    public Inspector()
    {
        InitializeComponent();
        RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL, VK_F);

        treeViewAll.AfterSelect += (s, el) =>
        {
            textBoxDetails.Clear();
            textBoxSuggestions.Clear();

            if (el.Node.Tag is AutomationElement element)
            {
                foreach (var prop in element.GetSupportedProperties())
                {
                    var line = string.Empty;
                    switch (prop.ProgrammaticName)
                    {
                        case "AutomationElementIdentifiers.ClassNameProperty":
                            line = $"ClassName: {element.Current.ClassName}";
                            break;
                        case "AutomationElementIdentifiers.NameProperty":
                            line = $"Name: {element.Current.Name}";
                            break;
                        case "AutomationElementIdentifiers.AutomationIdProperty":
                            line = $"AutomationId: {element.Current.AutomationId}";
                            break;
                        case "AutomationElementIdentifiers.ControlTypeProperty":
                            line = $"ControlType: {element.Current.LocalizedControlType}";
                            break;
                        case "AutomationElementIdentifiers.BoundingRectangleProperty":
                            var rect = element.GetCurrentPropertyValue(prop) as Rect?;
                            line = $"BoundingRectangle: {rect?.X}, {rect?.Y}, {rect?.Width}, {rect?.Height}";
                            break;
                        case "AutomationElementIdentifiers.ParentProperty":
                            var parent = element.GetCurrentPropertyValue(prop) as AutomationElement;
                            line = $"Parent: {parent?.Current.Name}";
                            break;
                        default:
                            continue;
                    }
                    textBoxDetails.AppendText(line + Environment.NewLine);
                }

                try
                {
                    StringBuilder suggestions = new();
                    var name = element.Current.Name;
                    var className = element.Current.ClassName;
                    var automationId = element.Current.AutomationId;

                    var window = GetParentWindow(element);
                    suggestions.AppendLine(
                        $".WindowControl(Name={window.Current.Name}, ClassName={window.Current.ClassName})"
                    );

                    static string Quote(string s) => s is null ? "''" : "\"" + s.Replace("\"", "\"\"") + "\"";
                    var locType = element.Current.LocalizedControlType?.ToLowerInvariant() ?? "";
                    if (!PythonControlMap.TryGetValue(locType, out var pyType))
                        pyType = "CustomControl";

                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(className) && !string.IsNullOrWhiteSpace(automationId))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(Name={Quote(name)}, ClassName={Quote(className)}, AutomationId={Quote(automationId)})"
                        );
                    }
                    else if(!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(className))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(Name={Quote(name)}, ClassName={Quote(className)})"
                        );
                    }
                    else if (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(automationId))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(Name={Quote(name)}, AutomationId={Quote(automationId)})"
                        );
                    }
                    else if (!string.IsNullOrEmpty(className) && !string.IsNullOrWhiteSpace(automationId))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(ClassName={Quote(className)}, AutomationId={Quote(automationId)})"
                        );
                    }
                    else if (!string.IsNullOrEmpty(name))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(Name={Quote(name)})"
                        );
                    }
                    else if (!string.IsNullOrEmpty(className))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(ClassName={Quote(className)})"
                        );
                    }
                    else if (!string.IsNullOrEmpty(automationId))
                    {
                        suggestions.AppendLine(
                            $".{pyType}(AutomationId={Quote(automationId)})"
                        );
                    }
                    else
                    {
                    }
                    textBoxSuggestions.Text = suggestions.ToString();
                }
                catch (Exception ex)
                {
                    textBoxSuggestions.Text = $"Error generating suggestions: {ex.Message}";
                }
            }
        };
    }

    private ContextMenuStrip contextMenu;

    private void SetupContextMenu()
    {
        contextMenu = new ContextMenuStrip();

        var menuClick = new ToolStripMenuItem("Click");
        menuClick.Click += (s, e) => PerformClickOnSelectedNode();

        var menuRightClick = new ToolStripMenuItem("Right Click");
        menuRightClick.Click += (s, e) => PerformRightClickOnSelectedNode();

        var menuType = new ToolStripMenuItem("Type Text");
        menuType.Click += (s, e) => PerformTypingOnSelectedNode();

        var menuDoubleClick = new ToolStripMenuItem("Double Click");
        menuDoubleClick.Click += (s, e) => PerformDoubleClickOnSelectedNode();

        contextMenu.Items.Add(menuClick);
        contextMenu.Items.Add(menuRightClick);
        contextMenu.Items.Add(menuDoubleClick);
        contextMenu.Items.Add(menuType);

        treeViewAll.NodeMouseClick += (s, e) =>
        {
            if (e.Button == MouseButtons.Right)
            {
                treeViewAll.SelectedNode = e.Node;
                contextMenu.Show(treeViewAll, e.Location);
            }
        };
    }

    // add these at the top of your class:
    [DllImport("user32.dll")]
    static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")]
    static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    const uint MOUSEEVENTF_LEFTUP = 0x0004;

    private void PerformClickOnSelectedNode()
    {
        if (treeViewAll.SelectedNode?.Tag is not AutomationElement element) return;
        var rect = element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty) as Rect?;
        if (rect.HasValue)
        {
            int x = (int)(rect.Value.X + rect.Value.Width / 2);
            int y = (int)(rect.Value.Y + rect.Value.Height / 2);
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
        }
    }

    private void PerformRightClickOnSelectedNode()
    {
        if (treeViewAll.SelectedNode?.Tag is not AutomationElement element) return;

        var rect = element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty) as Rect?;
        if (rect.HasValue)
        {
            int x = (int)(rect.Value.X + rect.Value.Width / 2);
            int y = (int)(rect.Value.Y + rect.Value.Height / 2);
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_RIGHTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_RIGHTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
        }
    }
    private void PerformDoubleClickOnSelectedNode()
    {
        if (treeViewAll.SelectedNode?.Tag is not AutomationElement element) return;

        var rect = element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty) as Rect?;
        if (rect.HasValue)
        {
            var x = (int)(rect.Value.X + rect.Value.Width / 2);
            var y = (int)(rect.Value.Y + rect.Value.Height / 2);
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
        }
    }

    private void PerformTypingOnSelectedNode()
    {
        if (treeViewAll.SelectedNode?.Tag is not AutomationElement element) return;

        var rect = element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty) as Rect?;
        if (rect.HasValue)
        {
            var x = (int)(rect.Value.X + rect.Value.Width / 2);
            var y = (int)(rect.Value.Y + rect.Value.Height / 2);
            
            var str = Microsoft.VisualBasic.Interaction.InputBox(
            "Please type the text you want to send to the selected element:",
            "User Input",
            ""
            );

            if (!string.IsNullOrEmpty(str))
            {
                SetCursorPos(x, y);
                mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)x, (uint)y, 0, UIntPtr.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, UIntPtr.Zero);
                SendKeys.SendWait(str);
            }
        }

    }

    private static TreeNode FindNodeByText(TreeNode root, string text)
    {
        if (root.Text.Contains(text, StringComparison.InvariantCultureIgnoreCase))
            return root;

        foreach (TreeNode child in root.Nodes)
        {
            var found = FindNodeByText(child, text);
            if (found != null)
                return found;
        }
        return null;
    }

    private void InitializeComponent()
    {
        splitContainer = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical
        };

        Load += (s, e) => splitContainer.SplitterDistance = ClientSize.Width * 3 / 5;

        var leftPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        leftPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        leftPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        searchBar = new TextBox
        {
            Dock = DockStyle.Top,
            PlaceholderText = "Search...",
            AutoSize = true
        };
        leftPanel.Controls.Add(searchBar, 0, 0);

        searchBar.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                var searchText = searchBar.Text.ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(searchText))
                    return;

                foreach (TreeNode node in treeViewAll.Nodes)
                {
                    var found = FindNodeByText(node, searchText);
                    if (found != null)
                    {
                        treeViewAll.SelectedNode = found;
                        treeViewAll.Focus();
                        found.EnsureVisible();
                        break;
                    }
                }
            }
        };

        treeViewAll = new TreeView
        {
            Dock = DockStyle.Fill,
            HideSelection = false
        };
        leftPanel.Controls.Add(treeViewAll);
        splitContainer.Panel1.Controls.Add(leftPanel);

        var rightPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

        textBoxDetails = new TextBox { Dock = DockStyle.Fill, Multiline = true };
        rightPanel.Controls.Add(textBoxDetails, 0, 0);

        textBoxSuggestions = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true
        };
        rightPanel.Controls.Add(textBoxSuggestions, 0, 1);

        splitContainer.Panel2.Controls.Add(rightPanel);

        Controls.Add(splitContainer);

        Text = "Inspector Tree";
        Width = 1000;
        Height = 700;

        SetupContextMenu();
    }


    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
        {
            if (GetCursorPos(out var pt))
            {
                var wp = new Point(pt.X, pt.Y);
                var el = AutomationElement.FromPoint(wp);
                ShowInfo(el);
            }
        }

        base.WndProc(ref m);
    }

    void ShowInfo(AutomationElement e)
    {
        if (e == null) return;

        AutomationElement window = GetParentWindow(e);
        if (window == null) return;

        // Build the automation tree for the parent window
        var fullRoot = BuildAutomationTree(window);

        // Setup the tree view
        treeViewAll.Nodes.Clear();
        var fullRootNode = CreateTreeNode(fullRoot);
        treeViewAll.Nodes.Add(fullRootNode);

        var selectedNode = FindNode(fullRootNode, e);
        if (selectedNode != null)
        {
            treeViewAll.Focus();
            treeViewAll.SelectedNode = selectedNode;
            selectedNode.EnsureVisible();
        }

    }

    private static TreeNode FindNode(TreeNode root, AutomationElement element)
    {
        if (root.Tag is AutomationElement el && el.Equals(element))
            return root;

        foreach (TreeNode child in root.Nodes)
        {
            var found = FindNode(child, element);
            if (found != null)
                return found;
        }

        return null;
    }

    private static TreeNode CreateTreeNode(TreeNode source)
    {
        var newNode = new TreeNode(source.Text)
        {
            Tag = source.Tag
        };
        foreach (TreeNode child in source.Nodes)
            newNode.Nodes.Add(CreateTreeNode(child));

        return newNode;
    }

    static TreeNode BuildAutomationTree(AutomationElement el)
    {
        var candidates = new[]
        {
            el.Current.Name,
            el.Current.LocalizedControlType,
            el.Current.ItemType,
            el.Current.AutomationId,
            el.Current.ClassName
        };

        var name = candidates.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));

        var rootNode = new TreeNode(name)
        {
            Tag = el
        };

        var children = el.FindAll(
            TreeScope.Children,
            System.Windows.Automation.Condition.TrueCondition);

        foreach (AutomationElement child in children)
            rootNode.Nodes.Add(BuildAutomationTree(child));

        return rootNode;
    }

    static AutomationElement GetParentWindow(AutomationElement e)
    {
        var walker = TreeWalker.ControlViewWalker;
        var p = walker.GetParent(e);
        while (p != null && p.Current.ControlType != ControlType.Window)
            p = walker.GetParent(p);
        return p;
    }
}


[StructLayout(LayoutKind.Sequential)]
public struct POINT
{
    public int X;
    public int Y;
}