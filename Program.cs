using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using System.Linq;

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

    public Inspector()
    {
        InitializeComponent();
        RegisterHotKey(Handle, HOTKEY_ID, MOD_CONTROL, VK_F);

        treeViewAll.AfterSelect += (s, el) =>
        {
            // Limpa o TextBox
            textBoxDetails.Clear();
            if (el.Node.Tag is AutomationElement element)
            {
                foreach (var prop in element.GetSupportedProperties())
                {

                    // var value = element.GetCurrentPropertyValue(prop) ?? "(null)";
                    // line = $"{prop.ProgrammaticName}: {value}";
                    
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
            }
        };
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

        Load += (s, e) => splitContainer.SplitterDistance = (int)(this.ClientSize.Width * 2/3);

        // TableLayoutPanel no lado esquerdo
        var leftPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        leftPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        leftPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        // search bar
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

        // Árvore completa
        treeViewAll = new TreeView
        {
            Dock = DockStyle.Fill,
            HideSelection = false
        };
        leftPanel.Controls.Add(treeViewAll);

        // adiciona o leftContainer ao painel esquerdo do split principal
        splitContainer.Panel1.Controls.Add(leftPanel);

        // Detalhes do elemento
        textBoxDetails = new TextBox { Dock = DockStyle.Fill, Multiline = true };
        splitContainer.Panel2.Controls.Add(textBoxDetails);

        Controls.Add(splitContainer);

        Text = "Inspector Tree";
        Width = 900;
        Height = 700;
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
        // Elemento inspecionado
        if (e == null)
        {
            Console.WriteLine("Elemento não encontrado.");
            return;
        }

        AutomationElement window = GetParentWindow(e);
        if (window == null) return;

        // Cria a árvore de renderização
        var fullRoot = BuildAutomationTree(window);

        // Exibe no treeViewAll
        treeViewAll.Nodes.Clear();
        var fullRootNode = CreateTreeNode(fullRoot);
        treeViewAll.Nodes.Add(fullRootNode);

        // Seleciona o elemento na árvore
        var selectedNode = FindNode(fullRootNode, e);
        if (selectedNode != null)
        {
            // Obter o foco da treeview
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
            Tag = source.Tag  // importante para o AfterSelect
        };
        foreach (TreeNode child in source.Nodes)
            newNode.Nodes.Add(CreateTreeNode(child));

        return newNode;
    }

    // monta recursivamente um ElementNode
    static TreeNode BuildAutomationTree(AutomationElement el)
    {
        // Se não houver Name, exibe o tipo de controle
        // tenta Name, depois LocalizedControlType, depois AutomationId, depois ClassName, senão "(unknown)"
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
            Tag = el // mantém o AutomationElement se precisar depois
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