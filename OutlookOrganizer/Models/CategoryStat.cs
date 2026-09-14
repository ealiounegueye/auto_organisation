using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace OutlookOrganizer.Models;

public sealed class CategoryStat : INotifyPropertyChanged
{
    private int _count;

    public required string Name { get; init; }
    public required Brush Color { get; init; }

    public int Count
    {
        get => _count;
        set
        {
            if (_count == value)
            {
                return;
            }

            _count = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(CountText));
        }
    }

    public string CountText => Count.ToString("N0");

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
