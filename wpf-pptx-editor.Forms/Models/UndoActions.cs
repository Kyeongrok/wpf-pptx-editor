using System.Collections.ObjectModel;

namespace wpf_pptx_editor.Forms.Models;

public interface IUndoableAction
{
    string Description { get; }
    void Undo();
}

public sealed class AddShapeAction(ObservableCollection<EditableShape> shapes, EditableShape added)
    : IUndoableAction
{
    public string Description => "도형 추가";
    public void Undo() => shapes.Remove(added);
}

public sealed class DeleteShapeAction(
    ObservableCollection<EditableShape> shapes, EditableShape removed, int index)
    : IUndoableAction
{
    public string Description => "도형 삭제";
    public void Undo() => shapes.Insert(Math.Min(index, shapes.Count), removed);
}

public sealed class MoveShapeAction(EditableShape shape, double oldX, double oldY)
    : IUndoableAction
{
    public string Description => "이동";
    public void Undo() { shape.X = oldX; shape.Y = oldY; }
}

public sealed class ChangePropertyAction<T>(Action<T> setter, T oldValue, string name)
    : IUndoableAction
{
    public string Description => $"{name} 변경";
    public void Undo() => setter(oldValue);
}
