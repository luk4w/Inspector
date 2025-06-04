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
    private nint lastHwnd;
    private static readonly Dictionary<int, string> ControlMapById = new()
{
    { ControlType.Button.Id,       "ButtonControl"     },
    { ControlType.Calendar.Id,     "CalendarControl"   },
    { ControlType.CheckBox.Id,     "CheckBoxControl"   },
    { ControlType.ComboBox.Id,     "ComboBoxControl"   },
    { ControlType.Custom.Id,       "CustomControl"     },
    { ControlType.DataGrid.Id,     "DataGridControl"   },
    { ControlType.DataItem.Id,     "DataItemControl"   },
    { ControlType.Document.Id,     "DocumentControl"   },
    { ControlType.Edit.Id,         "EditControl"       },
    { ControlType.Group.Id,        "GroupControl"      },
    { ControlType.Header.Id,       "HeaderControl"     },
    { ControlType.HeaderItem.Id,   "HeaderItemControl" },
    { ControlType.Hyperlink.Id,    "HyperlinkControl"  },
    { ControlType.Image.Id,        "ImageControl"      },
    { ControlType.List.Id,         "ListControl"       },
    { ControlType.ListItem.Id,     "ListItemControl"   },
    { ControlType.MenuBar.Id,      "MenuBarControl"    },
    { ControlType.Menu.Id,         "MenuControl"       },
    { ControlType.MenuItem.Id,     "MenuItemControl"   },
    { ControlType.Pane.Id,         "PaneControl"       },
    { ControlType.ProgressBar.Id,  "ProgressBarControl" },
    { ControlType.RadioButton.Id,  "RadioButtonControl" },
    { ControlType.ScrollBar.Id,    "ScrollBarControl"  },
    { ControlType.Slider.Id,       "SliderControl"     },
    { ControlType.Spinner.Id,      "SpinnerControl"    },
    { ControlType.SplitButton.Id,  "SplitButtonControl"},
    { ControlType.StatusBar.Id,    "StatusBarControl"  },
    { ControlType.Tab.Id,          "TabControl"        },
    { ControlType.TabItem.Id,      "TabItemControl"    },
    { ControlType.Table.Id,        "TableControl"      },
    { ControlType.Text.Id,         "TextControl"       },
    { ControlType.Thumb.Id,        "ThumbControl"      },
    { ControlType.TitleBar.Id,     "TitleBarControl"   },
    { ControlType.ToolBar.Id,      "ToolBarControl"    },
    { ControlType.ToolTip.Id,      "ToolTipControl"    },
    { ControlType.Tree.Id,         "TreeControl"       },
    { ControlType.TreeItem.Id,     "TreeItemControl"   },
    { ControlType.Window.Id,       "WindowControl"     }
};


    private bool _retrySelect = false;
    private void SelectElement(AutomationElement element)
    {
        try
        {         
            foreach (var prop in element.GetSupportedProperties())
            {
                var line = string.Empty;
                // list all properties of the element
                if (prop == AutomationElement.NameProperty)
                {
                    line = $"Name: {element.Current.Name}";
                }
                else if (prop == AutomationElement.ClassNameProperty)
                {
                    line = $"ClassName: {element.Current.ClassName}";
                }
                else if (prop == AutomationElement.AutomationIdProperty)
                {
                    line = $"AutomationId: {element.Current.AutomationId}";
                }
                else if (prop == AutomationElement.ControlTypeProperty)
                {
                    line = $"ControlType: {element.Current.ControlType.ProgrammaticName}";
                }
                else if (prop == AutomationElement.LocalizedControlTypeProperty)
                {
                    line = $"LocalizedControlType: {element.Current.LocalizedControlType}";
                }
                else
                {
                    continue; // Skip unsupported properties
                }
                textBoxDetails.AppendText(line + Environment.NewLine);
            }

            try
            {
                StringBuilder suggestions = new();
                var name = element.Current.Name;
                var className = element.Current.ClassName;
                var automationId = element.Current.AutomationId;
                var type = element.Current.ControlType.ProgrammaticName;

                var window = GetTopmostWindow(element);
                suggestions.AppendLine(
                    $".WindowControl(Name=\"{window.Current.Name}\", ClassName=\"{window.Current.ClassName}\")"
                );

                static string Quote(string s) => s is null ? "''" : "\"" + s.Replace("\"", "\"\"") + "\"";
                var locType = element.Current.LocalizedControlType?.ToLowerInvariant() ?? "";
                var controlId = element.Current.ControlType.Id;
                if (!ControlMapById.TryGetValue(controlId, out var pyType))
                {
                    pyType = "CustomControl";
                }

                if (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(className) && !string.IsNullOrWhiteSpace(automationId))
                {
                    suggestions.AppendLine(
                        $".{pyType}(Name={Quote(name)}, ClassName={Quote(className)}, AutomationId={Quote(automationId)})"
                    );
                }
                else if (!string.IsNullOrEmpty(name) && !string.IsNullOrWhiteSpace(className))
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
                    suggestions.AppendLine(
                        $".{pyType}()"
                    );
                }
                textBoxSuggestions.Text = suggestions.ToString();
            }
            catch (Exception ex)
            {
                textBoxSuggestions.Text = $"Error generating suggestions: {ex.Message}";
            }
        }
        catch (Exception)
        {
            ShowWindow(lastHwnd, SW_RESTORE);
            SetForegroundWindow(lastHwnd);

            if (_retrySelect == false)
            {
                _retrySelect = true;
                SelectElement(element);
                return;
            }
            else
            {
                textBoxDetails.AppendText("Application not found or element not available");
                _retrySelect = false;
                return;
            }
        }
    }

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
                SelectElement(element);
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
    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const int SW_RESTORE = 9;
    private static void ActiveParentWindow(AutomationElement element)
    {
        if (element == null) return;
        var window = GetTopmostWindow(element);
        if (window == null) return;
        var hwnd = new IntPtr(window.Current.NativeWindowHandle);
        if (hwnd == IntPtr.Zero) return;
        ShowWindow(hwnd, SW_RESTORE);
        SetForegroundWindow(hwnd);
    }

    private void PerformClickOnSelectedNode()
    {
        if (treeViewAll.SelectedNode?.Tag is not AutomationElement element) return;
        var rect = element.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty) as Rect?;
        if (rect.HasValue)
        {
            ActiveParentWindow(element);
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
            ActiveParentWindow(element);
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
            ActiveParentWindow(element);
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
            ActiveParentWindow(element);
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
        Width = 1200;
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

        AutomationElement window = GetTopmostWindow(e);
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
        lastHwnd = new IntPtr(window.Current.NativeWindowHandle);
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
        int ctrlId = el.Current.ControlType.Id;
        if (!ControlMapById.TryGetValue(ctrlId, out var controlType))
        {
            controlType = "CustomControl";
        }
        var candidates = new[]
        {
            el.Current.Name,
            el.Current.AutomationId,
            controlType
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

    private static AutomationElement GetTopmostWindow(AutomationElement e)
    {
        if (e == null) return null;
        var walker = TreeWalker.ControlViewWalker;
        AutomationElement topWindow = null;
        var current = e;

        while (current != null)
        {
            if (current.Current.ControlType == ControlType.Window)
                topWindow = current;
            current = walker.GetParent(current);
        }

        return topWindow;

    }

}


[StructLayout(LayoutKind.Sequential)]
public struct POINT
{
    public int X;
    public int Y;
}