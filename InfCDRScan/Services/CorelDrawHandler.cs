using Corel.Interop.VGCore;
using System;
using corel = Corel.Interop.VGCore;

namespace InfCDRScan.Services
{
    internal class CorelDrawHandler
    {
        private readonly corel.Application corelApp;
        private readonly FiltersManger filtersManger;

        private int pageID;

        public CorelDrawHandler(corel.Application app, FiltersManger filtersManger)
        {
            corelApp = app;
            this.filtersManger = filtersManger;
        }

        #region Разбирание коллекции по типам

        /// <summary>Сканирование коллекции shape типа</summary>
        /// <param name="sr">Коллекция shape типа</param>
        public void Scan(corel.ShapeRange sr, int pageID)
        {
            this.pageID = pageID;
            foreach (corel.Shape shape in sr)
            {
                if (shape.Type == corel.cdrShapeType.cdrGroupShape)
                    ProcessingOnGroupShape(shape);

                if (shape.PowerClip != null)
                    ProcessingOnPowerClipShape(shape);

                switch (shape.Type)
                {
                    case corel.cdrShapeType.cdrNoShape:
                        ProcessingOnNoShape(shape);
                        break;
                    case corel.cdrShapeType.cdrRectangleShape:
                        ProcessingOnRectangleShape(shape);
                        break;
                    case corel.cdrShapeType.cdrEllipseShape:
                        ProcessingOnEllipseShape(shape);
                        break;
                    case corel.cdrShapeType.cdrCurveShape:
                        ProcessingOnCurveShape(shape);
                        break;
                    case corel.cdrShapeType.cdrPolygonShape:
                        ProcessingOnPolygonShape(shape);
                        break;
                    case corel.cdrShapeType.cdrBitmapShape:
                        ProcessingOnBitmapShape(shape);
                        break;
                    case corel.cdrShapeType.cdrTextShape:
                        ProcessingOnTextShape(shape);
                        break;
                    case corel.cdrShapeType.cdrSelectionShape:
                        ProcessingOnSelectionShape(shape);
                        break;
                    case corel.cdrShapeType.cdrGuidelineShape:
                        ProcessingOnGuidelineShape(shape);
                        break;
                    case corel.cdrShapeType.cdrBlendGroupShape:
                        ProcessingOnBlendGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdrExtrudeGroupShape:
                        ProcessingOnExtrudeGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdrOLEObjectShape:
                        ProcessingOnOLEObjectShape(shape);
                        break;
                    case corel.cdrShapeType.cdrContourGroupShape:
                        ProcessingOnContourGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdrLinearDimensionShape:
                        ProcessingOnLinearDimensionShape(shape);
                        break;
                    case corel.cdrShapeType.cdrBevelGroupShape:
                        ProcessingOnBevelGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdrDropShadowGroupShape:
                        ProcessingOnDropShadowGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdr3DObjectShape:
                        ProcessingOn3DObjectShape(shape);
                        break;
                    case corel.cdrShapeType.cdrArtisticMediaGroupShape:
                        ProcessingOnArtisticMediaGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdrConnectorShape:
                        ProcessingOnConnectorShape(shape);
                        break;
                    case corel.cdrShapeType.cdrMeshFillShape:
                        ProcessingOnMeshFillShape(shape);
                        break;
                    case corel.cdrShapeType.cdrCustomShape:
                        ProcessingOnCustomShape(shape);
                        break;
                    case corel.cdrShapeType.cdrCustomEffectGroupShape:
                        ProcessingOnCustomEffectGroupShape(shape);
                        break;
                    case corel.cdrShapeType.cdrSymbolShape:
                        ProcessingOnSymbolShape(shape);
                        break;
                    case corel.cdrShapeType.cdrHTMLFormObjectShape:
                        ProcessingOnHTMLFormObjectShape(shape);
                        break;
                    case corel.cdrShapeType.cdrHTMLActiveObjectShape:
                        ProcessingOnHTMLActiveObjectShape(shape);
                        break;
                    case corel.cdrShapeType.cdrPerfectShape:
                        ProcessingOnPerfectShape(shape);
                        break;
                    case corel.cdrShapeType.cdrEPSShape:
                        ProcessingOnEPSShape(shape);
                        break;
                    default:
                        break;
                }
            }
        }

        #endregion

        #region методы обработки shape типа

        private void ProcessingOnGroupShape(corel.Shape shape) =>
            Scan(shape.Shapes.All(), pageID);

        private void ProcessingOnPowerClipShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            corel.cdrShapeType type = shape.Type;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.PowerClip,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(type), pageID),
                Icon = InfIconType.def
            });

            if (shape.Fill.Type != cdrFillType.cdrNoFill)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.PowerClipWithFill,
                    Description = string.Format("{0} | Page: {1}", GetShapeTypeName(type), pageID),
                    Icon = InfIconType.def
                });
            }

            Scan(shape.PowerClip.Shapes.All(), pageID);
        }

        private void ProcessingOnNoShape(corel.Shape shape)
        {

        }

        private void ProcessingOnRectangleShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.ShapeRectangle,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shape.Type), pageID),
                Icon = InfIconType.def
            });

            ScanFill(shape.Fill, shapeID, shape.Type);
            ScanOutline(shape.Outline, shapeID, shape.Type);
        }

        private void ProcessingOnEllipseShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.ShapeEllipse,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shape.Type), pageID),
                Icon = InfIconType.def
            });

            ScanFill(shape.Fill, shapeID, shape.Type);
            ScanOutline(shape.Outline, shapeID, shape.Type);
        }

        private void ProcessingOnCurveShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            int nodeCount = shape.Curve.Nodes.Count;

            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.ShapeCurve,
                Description = string.Format("{0} nodes | Page: {1}", nodeCount.ToString(), pageID.ToString()),
                Icon = InfIconType.def
            });

            if (nodeCount > 3000)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.ShapeNodesGreat,
                    Description = string.Format("{0} nodes | Page: {1}", nodeCount.ToString(), pageID.ToString()),
                    Icon = InfIconType.def
                });
            }

            ScanFill(shape.Fill, shapeID, shape.Type);
            ScanOutline(shape.Outline, shapeID, shape.Type);
        }

        private void ProcessingOnPolygonShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            string polygonType;
            InfIconType icon;

            switch (shape.Polygon.Type)
            {
                case cdrPolygonType.cdrPolygon:
                    icon = InfIconType.def;
                    polygonType = "Polygon";
                    break;
                case cdrPolygonType.cdrStar:
                    icon = InfIconType.def;
                    polygonType = "Complex Star";
                    break;
                case cdrPolygonType.cdrPolygonAsStar:
                    icon = InfIconType.def;
                    polygonType = "Star";
                    break;
                default:
                    icon = InfIconType.def;
                    polygonType = "Polygon";
                    break;
            }

            filtersManger.AddShape(new ShapeDataSheet(shape.StaticID, pageID)
            {
                FiltersType = InfFilters.ShapePolygon,
                Description = string.Format("{0} | Page: {1}", polygonType, pageID),
                Icon = icon
            });

            ScanFill(shape.Fill, shapeID, shape.Type);
            ScanOutline(shape.Outline, shapeID, shape.Type);
        }

        private void ProcessingOnBitmapShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            corel.Bitmap bitmap = shape.Bitmap;
            int resX = bitmap.ResolutionX;
            int resY = bitmap.ResolutionY;
            InfFilters modeFilter = InfFilters.Noname;
            string resolution = string.Format($"({resX}×{resY})");
            string commonDescription = string.Empty;
            string modeDescription = null;
            InfIconType icon = InfIconType.def;

            switch (bitmap.Mode)
            {
                case cdrImageType.cdrBlackAndWhiteImage:
                    commonDescription = "Black&White";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapBW;
                    break;
                case cdrImageType.cdr16ColorsImage:
                    commonDescription = "16 Color";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.Bitmap16Color;
                    break;
                case cdrImageType.cdrGrayscaleImage:
                    commonDescription = "Grayscale";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapGrayscale;
                    break;
                case cdrImageType.cdrPalettedImage:
                    commonDescription = "Paletted";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapPaletted;
                    break;
                case cdrImageType.cdrRGBColorImage:
                    commonDescription = "RGB Color";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapRGBColor;
                    break;
                case cdrImageType.cdrCMYKColorImage:
                    commonDescription = "CMYK Color";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapCMYKColor;
                    break;
                case cdrImageType.cdrDuotoneImage:
                    commonDescription = "Duotone";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapDuotone;
                    switch (bitmap.Duotone.Type)
                    {
                        case cdrDuotoneType.cdrMonotone:
                            modeDescription = "Monotone | Page";
                            icon = InfIconType.def;
                            break;
                        case cdrDuotoneType.cdrDuotone:
                            modeDescription = "Duotone | Page";
                            icon = InfIconType.def;
                            break;
                        case cdrDuotoneType.cdrTritone:
                            modeDescription = "Tritone | Page";
                            icon = InfIconType.def;
                            break;
                        case cdrDuotoneType.cdrQuadtone:
                            modeDescription = "Quadtone | Page";
                            icon = InfIconType.def;
                            break;
                        default:
                            break;
                    }

                    for (int i = 1; i <= bitmap.Duotone.InkCount; i++)
                    {
                        ProcessingColor(bitmap.Duotone.Inks[i].Color, shapeID, InfIconType.def);
                    }
                    break;
                case cdrImageType.cdrLABImage:
                    commonDescription = "LAB Color";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapLABColor;
                    break;
                case cdrImageType.cdrCMYKMultiChannelImage:
                    commonDescription = "CMYKMultiChannel";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapCMYKMultiChannel;
                    break;
                case cdrImageType.cdrRGBMultiChannelImage:
                    commonDescription = "RGBMultiChannel";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapRGBMultiChannel;
                    break;
                case cdrImageType.cdrSpotMultiChannelImage:
                    commonDescription = "SpotMultiChannel";
                    icon = InfIconType.def;
                    modeFilter = InfFilters.BitmapSpotMultiChannel;
                    break;
                default:
                    break;
            }

            //добавить в общий фильтр
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.BitmapCommon,
                Description = string.Format("{0} | {1} | Page: {2}", commonDescription, resolution, pageID),
                Icon = icon
            });

            //добавить в фильтр по цветовой модели
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = modeFilter,
                Description = string.Format("{0} | {1}: {2}", resolution, modeDescription ?? "Page", pageID),
                Icon = icon
            });

            //добавить в фильтр высокого разрешения
            if (resX > 320 || resY > 320)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.BitmapDPIGreat,
                    Description = string.Format("{0} | {1} | Page: {2}", commonDescription, resolution, pageID),
                    Icon = icon
                });
            }

            //добавить в фильтр непропорционального размера
            if (resX != resY)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.BitmapUnproportional,
                    Description = string.Format("{0} | {1} | Page: {2}", commonDescription, resolution, pageID),
                    Icon = icon
                });
            }

            //добавить в фильтр есть возможность кропа
            if (bitmap.Cropped)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.BitmapCropOn,
                    Description = string.Format("{0} | {1} | Page: {2}", commonDescription, resolution, pageID),
                    Icon = icon
                });
            }

            //добавить в фильтр есть маска прозрачности
            if (bitmap.Transparent)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.BitmapTransparent,
                    Description = string.Format("{0} | {1} | Page: {2}", commonDescription, resolution, pageID),
                    Icon = icon
                });
            }
        }

        private void ProcessingOnTextShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            string textType = string.Empty;
            InfIconType icon = InfIconType.def;
            corel.Text text = shape.Text;
            int count = text.Story.Characters.Count;
            switch (text.Type)
            {
                case cdrTextType.cdrArtisticText:
                    textType = "Artistic";
                    icon = InfIconType.def;
                    break;
                case cdrTextType.cdrParagraphText:
                    textType = "Paragraph";
                    icon = InfIconType.def;
                    break;
                case cdrTextType.cdrArtisticFittedText:
                    textType = "Artistic Fitted";
                    icon = InfIconType.def;
                    break;
                case cdrTextType.cdrParagraphFittedText:
                    textType = "Paragraph Fitted";
                    icon = InfIconType.def;
                    break;
                default:
                    break;
            }

            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.TextCommon,
                Description = string.Format("{0} | {1} chars | Page: {2}", textType, count, pageID),
                Icon = icon
            });

            if (!text.IsArtisticText && text.Overflow)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.TextOverflow,
                    Description = string.Format("{0} | {1} chars | Page: {2}", textType, count, pageID),
                    Icon = icon
                });
            }
        }

        private void ProcessingOnSelectionShape(corel.Shape shape)
        {

        }

        private void ProcessingOnGuidelineShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.Guidelines,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shape.Type), pageID),
                Icon = InfIconType.def
            });
        }

        private void ProcessingOnBlendGroupShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.EffectBlend,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shape.Type), pageID),
                Icon = InfIconType.def
            });
        }

        private void ProcessingOnExtrudeGroupShape(corel.Shape shape)
        {
            
        }

        private void ProcessingOnOLEObjectShape(corel.Shape shape)
        {
            
        }

        private void ProcessingOnContourGroupShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.EffectContour,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shape.Type), pageID),
                Icon = InfIconType.def
            });
            ProcessingColor(shape.Effect.Contour.FillColor, shapeID, InfIconType.def);
            ProcessingColor(shape.Effect.Contour.FillColorTo, shapeID, InfIconType.def);
            ProcessingColor(shape.Effect.Contour.OutlineColor, shapeID, InfIconType.def);
        }

        private void ProcessingOnLinearDimensionShape(corel.Shape shape)
        {
            int shapeID = shape.StaticID;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.LinearDimension,
                Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shape.Type), pageID),
                Icon = InfIconType.def
            });
        }

        private void ProcessingOnBevelGroupShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnDropShadowGroupShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOn3DObjectShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnArtisticMediaGroupShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnConnectorShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnMeshFillShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnCustomShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnCustomEffectGroupShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnSymbolShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnHTMLFormObjectShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnHTMLActiveObjectShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnPerfectShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        private void ProcessingOnEPSShape(corel.Shape shape)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region обработка заливки и обводки

        private void ScanFill(corel.Fill fill, int shapeID, corel.cdrShapeType shapeType)
        {
            switch (fill.Type)
            {
                case cdrFillType.cdrNoFill:
                    break;
                case cdrFillType.cdrUniformFill:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.FillUniform,
                        Description = string.Format("{0} | Page: {1}",GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    ProcessingColor(fill.UniformColor, shapeID, InfIconType.def);
                    break;
                case cdrFillType.cdrFountainFill:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.FillFountain,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    foreach (corel.FountainColor item in fill.Fountain.Colors)
                    {
                        ProcessingColor(item.Color, shapeID, InfIconType.def);
                    }
                    break;
                case cdrFillType.cdrPostscriptFill:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.FillPostscript,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    break;
                case cdrFillType.cdrTextureFill:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.FillTexture,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    break;
                case cdrFillType.cdrPatternFill:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.FillPattern,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    if (fill.Pattern.Type == cdrPatternFillType.cdrTwoColorPattern)
                    {
                        ProcessingColor(fill.Pattern.BackColor, shapeID, InfIconType.def);
                        ProcessingColor(fill.Pattern.FrontColor, shapeID, InfIconType.def);
                    }
                    break;
                case cdrFillType.cdrHatchFill:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.FillHatch,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    break;
                default:
                    break;
            }
        }

        private void ScanOutline(corel.Outline outline, int shapeID, corel.cdrShapeType shapeType)
        {
            switch (outline.Type)
            {
                case cdrOutlineType.cdrNoOutline:
                    break;
                case cdrOutlineType.cdrOutline:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.Outline,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    ProcessingColor(outline.Color, shapeID, InfIconType.def);
                    break;
                case cdrOutlineType.cdrEnhancedOutline:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.OutlineEnhanced,
                        Description = string.Format("{0} | Page: {1}", GetShapeTypeName(shapeType), pageID),
                        Icon = InfIconType.def
                    });
                    ProcessingColor(outline.Color, shapeID, InfIconType.def);
                    break;
                default:
                    break;
            }
        }

        #endregion

        #region обработка цвета
        /// <summary>
        /// Обработка цвета
        /// </summary>
        /// <param name="color">цвет</param>
        /// <param name="shapeID"></param>
        /// <param name="firstIcon">иконка типа объекта</param>
        private void ProcessingColor(corel.Color color, int shapeID, InfIconType firstIcon)
        {
            switch (color.Type)
            {
                case cdrColorType.cdrColorPantone:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorPantone,
                        Icon = firstIcon,
                        Description = GetPantoneColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorCMYK:
                    ProcessingCMYKColor(color, shapeID, firstIcon);
                    break;
                case cdrColorType.cdrColorCMY:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorCMY,
                        Icon = firstIcon,
                        Description = GetCMYColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorRGB:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorRGB,
                        Icon = firstIcon,
                        Description = GetRGBColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorHSB:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorHSB,
                        Icon = firstIcon,
                        Description = GetHSBColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorHLS:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorHLS,
                        Icon = firstIcon,
                        Description = GetHLSColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorBlackAndWhite:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorBW,
                        Icon = firstIcon,
                        Description = GetBWColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorGray:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorGray,
                        Icon = firstIcon,
                        Description = GetGrayColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorYIQ:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorYIQ,
                        Icon = firstIcon,
                        Description = GetYIQColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorLab:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorLab,
                        Icon = firstIcon,
                        Description = GetLabColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorPantoneHex:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorPantoneHEX,
                        Icon = firstIcon,
                        Description = GetPantoneHEXColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorRegistration:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorReg,
                        Icon = firstIcon,
                        Description = GetRegColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorSpot:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorSpot,
                        Icon = firstIcon,
                        Description = GetSpotColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorMixed:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorMixedColor,
                        Icon = firstIcon,
                        Description = GetColorMixedColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorUserInk:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorUserInk,
                        Icon = firstIcon,
                        Description = GetUserInkColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                case cdrColorType.cdrColorMultiChannel:
                    filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                    {
                        FiltersType = InfFilters.ColorMultiChannel,
                        Icon = firstIcon,
                        Description = GetMultiChannelColorName(color),
                        ColorHexValue = color.HexValue
                    });
                    break;
                default:
                    break;
            }
        }

        private void ProcessingCMYKColor(corel.Color color, int shapeID, InfIconType firstIcon)
        {
            int CMYKCyan = color.CMYKCyan;
            int CMYKMagenta = color.CMYKMagenta;
            int CMYKYellow = color.CMYKYellow;
            int CMYKBlack = color.CMYKBlack;
            filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
            {
                FiltersType = InfFilters.ColorCMYK,
                Description = GetCMYKColorName(CMYKCyan, CMYKMagenta, CMYKYellow, CMYKBlack),
                Icon = firstIcon,
                ColorHexValue = color.HexValue
            });

            if ((CMYKCyan + CMYKMagenta + CMYKYellow + CMYKBlack) > 300)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.CMYKTotalInkGreat,
                    Description = GetCMYKColorName(CMYKCyan, CMYKMagenta, CMYKYellow, CMYKBlack),
                    Icon = firstIcon,
                    ColorHexValue = color.HexValue
                });
            }

            if (CMYKCyan > 0 && CMYKMagenta > 0 && CMYKYellow > 0 && CMYKBlack > 0)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.CMYK400,
                    Description = GetCMYKColorName(CMYKCyan, CMYKMagenta, CMYKYellow, CMYKBlack),
                    Icon = firstIcon,
                    ColorHexValue = color.HexValue
                });
            }

            if ((CMYKCyan > 0 && CMYKCyan < 10) ||
                CMYKMagenta > 0 && CMYKMagenta < 10 ||
                CMYKYellow > 0 && CMYKYellow < 10 ||
                CMYKBlack > 0 && CMYKBlack < 10)
            {
                filtersManger.AddShape(new ShapeDataSheet(shapeID, pageID)
                {
                    FiltersType = InfFilters.CMYKMin10,
                    Description = GetCMYKColorName(CMYKCyan, CMYKMagenta, CMYKYellow, CMYKBlack),
                    Icon = firstIcon,
                    ColorHexValue = color.HexValue
                });
            }
        }

        #endregion

        #region Services

        #region текстовое представление типов

        #region текстовое представление цвета

        private string GetPantoneColorName(corel.Color color) =>
            string.Format("{0} | {1}%", color.SpotColorName, color.Tint);
        private string GetCMYKColorName(corel.Color color) =>
            GetCMYKColorName(color.CMYKCyan, color.CMYKMagenta, color.CMYKYellow, color.CMYKBlack);
        private string GetCMYKColorName(int CMYKCyan, int CMYKMagenta, int CMYKYellow, int CMYKBlack) =>
            string.Format("C:{0,-4}M:{1,-4}Y:{2,-4}K:{3}", CMYKCyan, CMYKMagenta, CMYKYellow, CMYKBlack);
        private string GetCMYColorName(corel.Color color) =>
            string.Format("C:{0,-4}M:{1,-4}Y:{2}", color.CMYCyan, color.CMYMagenta, color.CMYYellow);
        private string GetRGBColorName(corel.Color color) =>
            string.Format("R:{0,-4}G:{1,-4}B:{2}", color.RGBRed, color.RGBGreen, color.RGBBlue);
        private string GetHSBColorName(corel.Color color) =>
            string.Format("H:{0,-4}S:{1,-4}B:{2}", color.HSBHue, color.HSBSaturation, color.HSBBrightness);
        private string GetHLSColorName(corel.Color color) =>
            string.Format("H:{0,-4}L:{1,-4}S:{2}", color.HLSHue, color.HLSLightness, color.HLSSaturation);
        private string GetBWColorName(corel.Color color) =>
            string.Format("B:{0}", color.BW);
        private string GetGrayColorName(corel.Color color) =>
            string.Format("G:{0,-4}", color.Gray);
        private string GetYIQColorName(corel.Color color) =>
            string.Format("Y:{0,-4}I:{1,-4}Q:{2}", color.YIQLuminanceY, color.YIQChromaI, color.YIQChromaQ);
        private string GetLabColorName(corel.Color color) =>
            string.Format("L:{0,-4}a:{1,-4}b:{2}", color.LabLuminance, color.LabComponentA, color.LabComponentB);
        private string GetPantoneHEXColorName(corel.Color color) =>         //TODO нормальное имя 
            string.Format("{0} | {1}%", color.SpotColorName, color.Tint);
        private string GetRegColorName(corel.Color color) =>
            string.Format("{0}%", color.Tint);
        private string GetUserInkColorName(corel.Color color) =>            //TODO нормальное имя 
            string.Format("{0} | {1}%", color.Name, color.Tint);
        private string GetSpotColorName(corel.Color color) =>
            string.Format("{0} | {1}%", color.SpotColorName, color.Tint);
        private string GetMultiChannelColorName(corel.Color color) =>       //TODO нормальное имя 
            string.Format("{0}", color.Name);
        private string GetColorMixedColorName(corel.Color color) =>         //TODO нормальное имя 
            string.Format("{0}", color.Name);

        #endregion

        private string GetShapeTypeName(corel.cdrShapeType type)
        {
            string result = "Shape";

            switch (type)
            {
                case cdrShapeType.cdrNoShape:
                    break;
                case cdrShapeType.cdrRectangleShape:
                    result = "Rectangle";
                    break;
                case cdrShapeType.cdrEllipseShape:
                    result = "Ellipse";
                    break;
                case cdrShapeType.cdrCurveShape:
                    result = "Curve Shape";
                    break;
                case cdrShapeType.cdrPolygonShape:
                    result = "Polygon";
                    break;
                case cdrShapeType.cdrBitmapShape:
                    result = "Bitmap";
                    break;
                case cdrShapeType.cdrTextShape:
                    result = "Text";
                    break;
                case cdrShapeType.cdrGroupShape:
                    result = "Group";
                    break;
                case cdrShapeType.cdrSelectionShape:
                    break;
                case cdrShapeType.cdrGuidelineShape:
                    result = "Guideline";
                    break;
                case cdrShapeType.cdrBlendGroupShape:
                    result = "Blend Effect";
                    break;
                case cdrShapeType.cdrExtrudeGroupShape:
                    result = "Extrude Effect";
                    break;
                case cdrShapeType.cdrOLEObjectShape:
                    result = "OLE object";
                    break;
                case cdrShapeType.cdrContourGroupShape:
                    result = "Contour Effect";
                    break;
                case cdrShapeType.cdrLinearDimensionShape:
                    result = "Dimension";
                    break;
                case cdrShapeType.cdrBevelGroupShape:
                    result = "Bevel Effect";
                    break;
                case cdrShapeType.cdrDropShadowGroupShape:
                    result = "DropShadow Effect";
                    break;
                case cdrShapeType.cdr3DObjectShape:
                    result = "3D Object";
                    break;
                case cdrShapeType.cdrArtisticMediaGroupShape:
                    result = "Artistic Shape";
                    break;
                case cdrShapeType.cdrConnectorShape:
                    result = "Connector Shape";
                    break;
                case cdrShapeType.cdrMeshFillShape:
                    result = "MeshFill";
                    break;
                case cdrShapeType.cdrCustomShape:
                    break;
                case cdrShapeType.cdrCustomEffectGroupShape:
                    break;
                case cdrShapeType.cdrSymbolShape:
                    result = "Symbol Shape";
                    break;
                case cdrShapeType.cdrHTMLFormObjectShape:
                    break;
                case cdrShapeType.cdrHTMLActiveObjectShape:
                    break;
                case cdrShapeType.cdrPerfectShape:
                    result = "Perfect Shape";
                    break;
                case cdrShapeType.cdrEPSShape:
                    break;
                default:
                    break;
            }

            return result;
        }

        #endregion

        #endregion
    }
}
