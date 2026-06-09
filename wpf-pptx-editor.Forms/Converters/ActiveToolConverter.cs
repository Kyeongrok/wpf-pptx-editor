using System.Globalization;
using System.Windows.Data;
using wpf_pptx_editor.Forms.Models;

namespace wpf_pptx_editor.Forms.Converters;

public class ActiveToolConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values.Length < 2 || values[0] is not DrawingTool tool || values[1] is not string param)
            return false;
        return tool.ToString() == param;
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        => throw new NotSupportedException();
}
