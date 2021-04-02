using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using InfCDRScan.Models.Filters;
using InfCDRScan.Models.Shapes;
using corel = Corel.Interop.VGCore;

namespace InfCDRScan.Services
{
    internal enum InfFilterGroup
    {
        [Description("Common")]
        Common,
        [Description("Simple Shape")]
        Shape,
        [Description("Text")]
        Text,
        [Description("Bitmap")]
        Bitmap,
        [Description("PowerClip")]
        PowerClip,
        [Description("Color Palette")]
        ColorType,
        [Description("Prepress")]
        Prepress,
        [Description("Fill")]
        Fill,
        [Description("Outline")]
        Outline,
        [Description("Effect")]
        Effect,
        [Description("Special Shape")]
        Special,
        [Description("Services")]
        Services,
    }

    internal enum InfFilters
    {
        [Description("Noname"), Icon(InfIconType.def)]
        Noname,

        //простые векторные формы
        [Description("Curve Shape"), Icon(InfIconType.def), GroupName(InfFilterGroup.Shape)]
        ShapeCurve,
        [Description("Rectangle"), Icon(InfIconType.def), GroupName(InfFilterGroup.Shape)]
        ShapeRectangle,
        [Description("Ellipse"), Icon(InfIconType.def), GroupName(InfFilterGroup.Shape)]
        ShapeEllipse,
        [Description("Polygon"), Icon(InfIconType.def), GroupName(InfFilterGroup.Shape)]
        ShapePolygon,
        [Description("Shape Nodes > 3000"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ShapeNodesGreat,

        //текст
        [Description("Text"), Icon(InfIconType.def), GroupName(InfFilterGroup.Text)]
        TextCommon,
        [Description("Text Overflow"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        TextOverflow,
        [Description("Different Text Fill"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        TextDifferentFill,

        //bitmaps
        [Description("Bitmap"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapCommon,
        [Description("Bitmap Black&White"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapBW,
        [Description("Bitmap Color 16"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        Bitmap16Color,
        [Description("Bitmap Grayscale"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapGrayscale,
        [Description("Bitmap Paletted"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapPaletted,
        [Description("Bitmap Color RGB"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapRGBColor,
        [Description("Bitmap Color CMYK"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapCMYKColor,
        [Description("Bitmap Duotone"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapDuotone,
        [Description("Bitmap Color LAB"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapLABColor,
        [Description("Bitmap CMYKMultiChannel"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapCMYKMultiChannel,
        [Description("Bitmap RGBMultiChannel"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapRGBMultiChannel,
        [Description("Bitmap SpotMultiChannel"), Icon(InfIconType.def), GroupName(InfFilterGroup.Bitmap)]
        BitmapSpotMultiChannel,

        //powerclip
        [Description("PowerClip"), Icon(InfIconType.def), GroupName(InfFilterGroup.PowerClip)]
        PowerClip,
        [Description("PowerClip with Fill"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        PowerClipWithFill,

        //простой цвет
        [Description("Color Pantone"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorPantone,
        [Description("Color CMYK"), Icon(InfIconType.CMYKColorModel), GroupName(InfFilterGroup.ColorType)]
        ColorCMYK,
        [Description("Color CMY"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorCMY,
        [Description("Color RGB"), Icon(InfIconType.RGBColorModel), GroupName(InfFilterGroup.ColorType)]
        ColorRGB,
        [Description("Color HSB"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorHSB,
        [Description("Color HLS"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorHLS,
        [Description("Color B&W"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorBW,
        [Description("Color Gray"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorGray,
        [Description("Color YIQ"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorYIQ,
        [Description("Color Lab"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorLab,
        [Description("Color Pantone HEX"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorPantoneHEX,
        [Description("Color Registration"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorReg,
        [Description("Color User ink"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorUserInk,
        [Description("Color Spot"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorSpot,
        [Description("Color Multi-channel"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorMultiChannel,
        [Description("Color Mixed"), Icon(InfIconType.def), GroupName(InfFilterGroup.ColorType)]
        ColorMixedColor,

        //препресс
        [Description("Bitmap Res. > 320dpi"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        BitmapDPIGreat,
        [Description("Bitmap Unproportional"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        BitmapUnproportional,
        [Description("Bitmap Crop On"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        BitmapCropOn,
        [Description("Bitmap Transparent"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        BitmapTransparent,
        [Description("Bitmap Overprint"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        BitmapOverprint,
        [Description("Total ink > 300%"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        CMYKTotalInkGreat,
        [Description("CMYK 400"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        CMYK400,
        [Description("Color control (min 10)"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        CMYKMin10,
        [Description("Color White Overprint"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ObjectOverprintWhite,
        [Description("Object Hidden"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ObjectHide,
        [Description("Object Locked"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ObjectLock,
        [Description("Group of 1"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ObjectGroupOne,

        //заливки
        [Description("Fill Uniform"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillUniform,
        [Description("Fill Fountain"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillFountain,
        [Description("Fill Postscript"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillPostscript,
        [Description("Fill Texture"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillTexture,
        [Description("Fill Pattern"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillPattern,
        [Description("Fill Hatch"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillHatch,
        [Description("Fill Mesh"), Icon(InfIconType.def), GroupName(InfFilterGroup.Fill)]
        FillMesh,
        [Description("Fill Overprint"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ObjectOverprintFill,

        //обводки
        [Description("Outline"), Icon(InfIconType.def), GroupName(InfFilterGroup.Outline)]
        Outline,
        [Description("Outline Enhanced"), Icon(InfIconType.def), GroupName(InfFilterGroup.Outline)]
        OutlineEnhanced,
        [Description("Outline Overprint"), Icon(InfIconType.def), GroupName(InfFilterGroup.Prepress)]
        ObjectOverprintOutline,

        //эфекты шейпов
        [Description("Blend"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectBlend,
        [Description("Extrude"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectExtrude,
        [Description("Envelope"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectEnvelope,
        [Description("TextOnPathEffect"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectTextOnPath,
        [Description("ControlPathEffect"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectControlPath,
        [Description("DropShadow"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectDropShadow,
        [Description("Contour"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectContour,
        [Description("Distortion"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectDistortion,
        [Description("Perspective"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectPerspective,
        [Description("Lens"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectLens,
        [Description("Custom Effect"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectCustom,
        [Description("Bevel"), Icon(InfIconType.def), GroupName(InfFilterGroup.Effect)]
        EffectBevel,

        //разные шейпы
        [Description("ArtisticMedia"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        EffectArtisticMedia,
        [Description("3D Object"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        Effect3DObject,
        [Description("HTML Form"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ObjectHTMLForm,
        [Description("EPS"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ShapeEPS,
        [Description("Custom Shape"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ShapeCustom,
        [Description("Perfect Shape"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ShapePerfect,
        [Description("OLE Shape"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ShapeOLE,
        [Description("Connector Shape"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ShapeConnector,
        [Description("Symbol Shape"), Icon(InfIconType.def), GroupName(InfFilterGroup.Special)]
        ShapeSymbol,

        //вспомогательные шейпы
        [Description("Guidelines"), Icon(InfIconType.def), GroupName(InfFilterGroup.Services)]
        ObjectGuidelines,
        [Description("Linear Dimension"), Icon(InfIconType.def), GroupName(InfFilterGroup.Services)]
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
                    GroupName = GetGroupName(item),
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
            if (attributes.Length > 0)
                return ((DescriptionAttribute)attributes[0]).Value;
            else
                return value.ToString();
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

        /// <summary>
        /// Возвращает значение атрибута GroupName из перечисления
        /// </summary>
        /// <param name="value">Элемент типа перечисления</param>
        /// <returns></returns>
        private string GetGroupName(Enum value)
        {
            Type type = value.GetType();
            FieldInfo fieldInfo = type.GetField(value.ToString());
            var attributes = fieldInfo.GetCustomAttributes(typeof(GroupNameAttribute), false);
            if (attributes.Length > 0)
                return GetDescription(((GroupNameAttribute)attributes[0]).Value);
            else
                return GetDescription(InfFilterGroup.Common);
        }

        #endregion
    }
}
