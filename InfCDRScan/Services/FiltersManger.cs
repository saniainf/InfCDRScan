using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using InfCDRScan.Models.Filters;
using InfCDRScan.Models.Shapes;
using corel = Corel.Interop.VGCore;

namespace InfCDRScan.Services
{
    internal enum InfFilters
    {
        [Description("Noname"), Icon(InfIconType.def)]
        Noname,

        //простые векторные формы
        [Description("Curve Shape"), Icon(InfIconType.def)]
        ShapeCurve,
        [Description("Rectangle"), Icon(InfIconType.def)]
        ShapeRectangle,
        [Description("Ellipse"), Icon(InfIconType.def)]
        ShapeEllipse,
        [Description("Polygon"), Icon(InfIconType.def)]
        ShapePolygon,
        [Description("Shape Nodes > 3000"), Icon(InfIconType.def)]
        ShapeNodesGreat,

        //текст
        [Description("Text"), Icon(InfIconType.def)]
        TextCommon,
        [Description("Overflow Text"), Icon(InfIconType.def)]
        TextOverflow,
        [Description("Different Text Fill"), Icon(InfIconType.def)]
        TextDifferentFill,

        //bitmaps
        [Description("Bitmap"), Icon(InfIconType.def)]
        BitmapCommon,
        [Description("Bitmap Res. > 320dpi"), Icon(InfIconType.def)]
        BitmapDPIGreat,
        [Description("Bitmap Unproportional"), Icon(InfIconType.def)]
        BitmapUnproportional,
        [Description("Bitmap Crop On"), Icon(InfIconType.def)]
        BitmapCropOn,
        [Description("Bitmap Transparent"), Icon(InfIconType.def)]
        BitmapTransparent,
        [Description("Bitmap Black&White"), Icon(InfIconType.def)]
        BitmapBW,
        [Description("Bitmap 16 color"), Icon(InfIconType.def)]
        Bitmap16Color,
        [Description("Bitmap Grayscale"), Icon(InfIconType.def)]
        BitmapGrayscale,
        [Description("Bitmap Paletted"), Icon(InfIconType.def)]
        BitmapPaletted,
        [Description("Bitmap RGB Color"), Icon(InfIconType.def)]
        BitmapRGBColor,
        [Description("Bitmap CMYK Color"), Icon(InfIconType.def)]
        BitmapCMYKColor,
        [Description("Bitmap Duotone"), Icon(InfIconType.def)]
        BitmapDuotone,
        [Description("Bitmap LAB Color"), Icon(InfIconType.def)]
        BitmapLABColor,
        [Description("Bitmap CMYKMultiChannel"), Icon(InfIconType.def)]
        BitmapCMYKMultiChannel,
        [Description("Bitmap RGBMultiChannel"), Icon(InfIconType.def)]
        BitmapRGBMultiChannel,
        [Description("Bitmap SpotMultiChannel"), Icon(InfIconType.def)]
        BitmapSpotMultiChannel,

        //powerclip
        [Description("PowerClip"), Icon(InfIconType.def)]
        PowerClip,
        [Description("PowerClip with Fill"), Icon(InfIconType.def)]
        PowerClipWithFill,

        //простой цвет
        [Description("Pantone Color"), Icon(InfIconType.def)]
        ColorPantone,
        [Description("CMYK Color"), Icon(InfIconType.CMYKColorModel)]
        ColorCMYK,
        [Description("CMY Color"), Icon(InfIconType.def)]
        ColorCMY,
        [Description("RGB Color"), Icon(InfIconType.RGBColorModel)]
        ColorRGB,
        [Description("HSB Color"), Icon(InfIconType.def)]
        ColorHSB,
        [Description("HLS Color"), Icon(InfIconType.def)]
        ColorHLS,
        [Description("B&W Color"), Icon(InfIconType.def)]
        ColorBW,
        [Description("Gray Color"), Icon(InfIconType.def)]
        ColorGray,
        [Description("YIQ Color"), Icon(InfIconType.def)]
        ColorYIQ,
        [Description("Lab Color"), Icon(InfIconType.def)]
        ColorLab,
        [Description("Pantone HEX Color"), Icon(InfIconType.def)]
        ColorPantoneHEX,
        [Description("Registration Color"), Icon(InfIconType.def)]
        ColorReg,
        [Description("User ink Color"), Icon(InfIconType.def)]
        ColorUserInk,
        [Description("Spot Color"), Icon(InfIconType.def)]
        ColorSpot,
        [Description("Multi-channel Color"), Icon(InfIconType.def)]
        ColorMultiChannel,
        [Description("Mixed Color"), Icon(InfIconType.def)]
        ColorMixedColor,

        //фильтры для cmyk
        [Description("Total ink > 300%"), Icon(InfIconType.def)]
        CMYKTotalInkGreat,
        [Description("CMYK 400"), Icon(InfIconType.def)]
        CMYK400,
        [Description("Color control (min 10)"), Icon(InfIconType.def)]
        CMYKMin10,

        //заливки
        [Description("Uniform Fill"), Icon(InfIconType.def)]
        FillUniform,
        [Description("Fountain Fill"), Icon(InfIconType.def)]
        FillFountain,
        [Description("Postscript Fill"), Icon(InfIconType.def)]
        FillPostscript,
        [Description("Texture Fill"), Icon(InfIconType.def)]
        FillTexture,
        [Description("Pattern Fill"), Icon(InfIconType.def)]
        FillPattern,
        [Description("Hatch Fill"), Icon(InfIconType.def)]
        FillHatch,
        [Description("Mesh Fill"), Icon(InfIconType.def)]
        FillMesh,

        //обводки
        [Description("Outline"), Icon(InfIconType.def)]
        Outline,
        [Description("Enhanced Outline"), Icon(InfIconType.def)]
        OutlineEnhanced,

        //эфекты шейпов
        [Description("Blend"), Icon(InfIconType.def)]
        EffectBlend,
        [Description("Extrude"), Icon(InfIconType.def)]
        EffectExtrude,
        [Description("Envelope"), Icon(InfIconType.def)]
        EffectEnvelope,
        [Description("TextOnPathEffect"), Icon(InfIconType.def)]
        EffectTextOnPath,
        [Description("ControlPathEffect"), Icon(InfIconType.def)]
        EffectControlPath,
        [Description("DropShadow"), Icon(InfIconType.def)]
        EffectDropShadow,
        [Description("Contour"), Icon(InfIconType.def)]
        EffectContour,
        [Description("Distortion"), Icon(InfIconType.def)]
        EffectDistortion,
        [Description("Perspective"), Icon(InfIconType.def)]
        EffectPerspective,
        [Description("Lens"), Icon(InfIconType.def)]
        EffectLens,
        [Description("Custom Effect"), Icon(InfIconType.def)]
        EffectCustom,
        [Description("Bevel"), Icon(InfIconType.def)]
        EffectBevel,

        //разные шейпы
        [Description("ArtisticMedia"), Icon(InfIconType.def)]
        EffectArtisticMedia,
        [Description("3D Object"), Icon(InfIconType.def)]
        Effect3DObject,
        [Description("HTML Form"), Icon(InfIconType.def)]
        ObjectHTMLForm,
        [Description("EPS"), Icon(InfIconType.def)]
        ShapeEPS,
        [Description("Custom Shape"), Icon(InfIconType.def)]
        ShapeCustom,
        [Description("Perfect Shape"), Icon(InfIconType.def)]
        ShapePerfect,
        [Description("OLE Shape"), Icon(InfIconType.def)]
        ShapeOLE,
        [Description("Connector Shape"), Icon(InfIconType.def)]
        ShapeConnector,
        [Description("Symbol Shape"), Icon(InfIconType.def)]
        ShapeSymbol,

        //вспомогательные шейпы
        [Description("Guidelines"), Icon(InfIconType.def)]
        ObjectGuidelines,
        [Description("Linear Dimension"), Icon(InfIconType.def)]
        ObjectLinearDimension,
    }

    internal class FiltersManger
    {
        private readonly Dictionary<InfFilters, FilterModel> filters;
        private readonly int docID;

        public FiltersManger(int docID)
        {
            this.docID = docID;
            filters = new Dictionary<InfFilters, FilterModel>();
            CreateFilters();
        }

        public List<FilterModel> GetFilters()
        {
            return filters.Values
                .Where(f => f.Shapes.Count > 0)
                .ToList();
        }

        public void AddShape(ShapeDataSheet data)
        {
            filters[data.FiltersType].Shapes.Add(CreateShape(data));
        }

        #region Services методы

        private void CreateFilters()
        {
            foreach (InfFilters item in Enum.GetValues(typeof(InfFilters)))
            {
                var newFilter = new FilterModel()
                {
                    Icon = GetIcon(item),
                    Description = GetDescription(item),
                    Shapes = new List<ShapeModel>()
                };
                filters.Add(item, newFilter);
            }
        }

        private ShapeModel CreateShape(ShapeDataSheet data)
        {
            return new ShapeModel(data.ShapeID, data.PageID, docID)
            {
                Description = data.Description,
                Icon = data.Icon,
                SecondIconVisibility = !string.IsNullOrEmpty(data.ColorHexValue),
                ColorHexValue = data.ColorHexValue
            };
        }

        /// <summary>
        /// Возвращает значение атрибута FilterDescription из перечисления
        /// </summary>
        /// <param name="value">Элемент типа перечисления</param>
        /// <returns></returns>
        private string GetDescription(Enum value)
        {
            Type type = value.GetType();
            FieldInfo fieldInfo = type.GetField(value.ToString());
            var attributes = fieldInfo.GetCustomAttributes(typeof(DescriptionAttribute), false);
            var attribute = (DescriptionAttribute)attributes[0];
            return attribute.Value;
        }

        /// <summary>
        /// Возвращает значение атрибута FilterIcon из перечисления
        /// </summary>
        /// <param name="value">Элемент типа перечисления</param>
        /// <returns></returns>
        private InfIconType GetIcon(Enum value)
        {
            Type type = value.GetType();
            FieldInfo fieldInfo = type.GetField(value.ToString());
            var attributes = fieldInfo.GetCustomAttributes(typeof(IconAttribute), false);
            if (attributes.Length > 0)
                return ((IconAttribute)attributes[0]).Value;
            else
                return InfIconType.None;
        }

        #endregion
    }
}
