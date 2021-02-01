using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace InfCDRScan.Services
{
    internal enum InfIconType
    {
        None = 0,
        def = 1,
        CMYKColorModel = 2,
        RGBColorModel = 3
    }

    class InfIcon : Image
    {
        public static readonly DependencyProperty IconProperty =
            DependencyProperty.Register("Icon", typeof(InfIconType), typeof(InfIcon), new PropertyMetadata(InfIconType.None, OnIconPropertyChanged));

        public InfIconType Icon
        {
            get { return (InfIconType)GetValue(IconProperty); }
            set
            {
                SetValue(IconProperty, value);
            }
        }

        private static void OnIconPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (!(d is InfIcon infIcon)) return;

            string name = Enum.GetName(typeof(InfIconType), infIcon.Icon);
            Uri uri = new Uri(string.Format("pack://application:,,,/InfCDRScan;component/Resources/Icon/{0}.gif", name));
            BitmapImage bitmap = new BitmapImage(uri);
            infIcon.SetValue(SourceProperty, bitmap);
        }
    }
}
