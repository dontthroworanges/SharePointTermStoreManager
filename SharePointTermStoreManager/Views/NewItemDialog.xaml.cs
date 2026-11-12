using Microsoft.UI.Xaml.Controls;
using SharePointTermStoreManager.Helpers;
using SharePointTermStoreManager.Models;

namespace SharePointTermStoreManager.Views;

/// <summary>Asks for the name and description of a new term group, term set or term. Doesn't create anything.</summary>
public sealed partial class NewItemDialog : ContentDialog
{
    private readonly TermNode? _parent;
    private readonly IReadOnlyCollection<TermNode> _siblings;

    /// <param name="siblings">Items already at that level, to catch duplicate names before SharePoint does.</param>
    public NewItemDialog(TermNodeKind kind, TermNode? parent, IReadOnlyCollection<TermNode> siblings)
    {
        InitializeComponent();
        _parent = parent;
        _siblings = siblings;

        Title = kind switch
        {
            TermNodeKind.Group => "New term group",
            TermNodeKind.TermSet => "New term set",
            _ => parent?.Kind == TermNodeKind.Term ? "New child term" : "New term",
        };
        LocationText.Text = parent == null ? "In the term store." : $"In {parent.Path}";
        NameBox.Header = kind == TermNodeKind.Term ? "Default label" : "Name";
        Opened += (_, _) => NameBox.Focus(Microsoft.UI.Xaml.FocusState.Programmatic);
    }

    /// <summary>The normalized name, once the dialog is accepted.</summary>
    public string ItemName { get; private set; } = "";

    public string ItemDescription { get; private set; } = "";

    private void ContentDialog_PrimaryButtonClick(ContentDialog sender, ContentDialogButtonClickEventArgs args)
    {
        var name = TermNames.Normalize(NameBox.Text);
        var description = (DescriptionBox.Text ?? "").Trim().Replace("\r\n", "\n").Replace('\r', '\n');

        var error = TermNames.Validate(name, "name") ?? TermNames.ValidateDescription(description);
        if (error == null && _siblings.Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)))
            error = _parent == null
                ? $"There's already a term group named “{name}”."
                : $"“{_parent.Name}” already contains “{name}”.";

        if (error != null)
        {
            ErrorBar.Message = error;
            ErrorBar.IsOpen = true;
            args.Cancel = true;
            return;
        }

        ItemName = name;
        ItemDescription = description;
    }
}
