using Microsoft.UI.Input;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Input;
using Windows.System;

namespace SharePointTermStoreManager.Views;

/// <summary>
/// A vertical splitter between two panes: drag it (left-right resize cursor), or focus it and press the Left and
/// Right arrow keys; double-click to reset. It only reports movement; the owner resizes the columns.
/// </summary>
public sealed partial class PaneSplitter : UserControl
{
    /// <summary>How far one arrow key press moves it (Shift: further).</summary>
    private const double KeyStep = 16, LargeKeyStep = 64;

    private double _startX;
    private bool _dragging;
    private bool _moved;
    private bool _pointerOver;

    public PaneSplitter()
    {
        InitializeComponent();
        ProtectedCursor = InputSystemCursor.Create(InputSystemCursorShape.SizeWestEast);
        GotFocus += (_, _) => UpdateGrip();
        LostFocus += (_, _) => UpdateGrip();
    }

    /// <summary>A drag started: remember the current size.</summary>
    public event EventHandler? DragStarted;

    /// <summary>Horizontal distance from where the drag started, in pixels (positive: to the right).</summary>
    public event EventHandler<double>? DragDelta;

    /// <summary>A drag that moved it ended, or an arrow key moved it: save the size.</summary>
    public event EventHandler? DragCompleted;

    /// <summary>An arrow key moved it by this many pixels.</summary>
    public event EventHandler<double>? Nudged;

    /// <summary>Double-clicked: go back to the default size.</summary>
    public event EventHandler? ResetRequested;

    protected override void OnPointerEntered(PointerRoutedEventArgs e)
    {
        _pointerOver = true;
        UpdateGrip();
    }

    protected override void OnPointerExited(PointerRoutedEventArgs e)
    {
        _pointerOver = false;
        UpdateGrip();
    }

    protected override void OnPointerPressed(PointerRoutedEventArgs e)
    {
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        if (!CapturePointer(e.Pointer)) return;
        _dragging = true;
        _moved = false;
        _startX = e.GetCurrentPoint(null).Position.X;
        DragStarted?.Invoke(this, EventArgs.Empty);
        UpdateGrip();
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerRoutedEventArgs e)
    {
        if (!_dragging) return;
        double delta = e.GetCurrentPoint(null).Position.X - _startX;
        if (delta == 0 && !_moved) return;
        _moved = true;
        DragDelta?.Invoke(this, delta);
        e.Handled = true;
    }

    protected override void OnPointerReleased(PointerRoutedEventArgs e)
    {
        if (!_dragging) return;
        ReleasePointerCapture(e.Pointer);
        EndDrag();
        e.Handled = true;
    }

    protected override void OnPointerCaptureLost(PointerRoutedEventArgs e) => EndDrag();

    protected override void OnDoubleTapped(DoubleTappedRoutedEventArgs e)
    {
        ResetRequested?.Invoke(this, EventArgs.Empty);
        e.Handled = true;
    }

    protected override void OnKeyDown(KeyRoutedEventArgs e)
    {
        double step = InputKeyboardSource.GetKeyStateForCurrentThread(VirtualKey.Shift)
            .HasFlag(Windows.UI.Core.CoreVirtualKeyStates.Down) ? LargeKeyStep : KeyStep;
        double? delta = e.Key switch
        {
            VirtualKey.Left => -step,
            VirtualKey.Right => step,
            _ => null,
        };
        if (delta == null) return;
        Nudged?.Invoke(this, delta.Value);
        DragCompleted?.Invoke(this, EventArgs.Empty);
        e.Handled = true;
    }

    /// <summary>Ends a drag; only a press that actually moved counts (the clicks of a double-click don't).</summary>
    private void EndDrag()
    {
        if (!_dragging) return;
        _dragging = false;
        if (_moved) DragCompleted?.Invoke(this, EventArgs.Empty);
        UpdateGrip();
    }

    private void UpdateGrip()
    {
        bool active = _dragging || _pointerOver || FocusState == FocusState.Keyboard;
        VisualStateManager.GoToState(this, active ? "Active" : "Idle", true);
    }
}
