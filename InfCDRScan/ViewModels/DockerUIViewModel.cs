using System.Windows.Input;
using corel = Corel.Interop.VGCore;
using InfCDRScan.Models.Filters;
using InfCDRScan.Models.Shapes;
using System.Linq;
using System.Windows.Data;
using System.ComponentModel;
using System.Collections.Generic;
using System.Diagnostics;
using InfCDRScan.ViewModels.Base;
using InfCDRScan.Services;
using InfCDRScan.Infrastructure.Commands;

namespace InfCDRScan.ViewModels
{
    internal class DockerUIViewModel : ViewModel
    {
        #region Поля

        private readonly corel.Application corelApp;

        private List<FilterModel> filters;

        private Stopwatch stopwatch;

        #endregion

        #region Свойства

        #region CollectionViewFilters : Представление коллекции фильтров

        private readonly CollectionViewSource collectionViewFilters = new CollectionViewSource();

        public ICollectionView CollectionViewFilters => collectionViewFilters?.View;

        #endregion

        #region CollectionViewShapes : Представление коллекции шейпов

        private readonly CollectionViewSource collectionViewShapes = new CollectionViewSource();

        public ICollectionView CollectionViewShapes => collectionViewShapes?.View;

        #endregion

        #region SelectedFilter : Выбранный фильтр
        /// <summary> выбранный фильтр </summary>
        private FilterModel selectedFilter;
        /// <summary> выбранный фильтр </summary>
        public FilterModel SelectedFilter
        {
            get => selectedFilter;
            set
            {
                if (!Set(ref selectedFilter, value)) return;
                if (selectedFilter == null) return;
                collectionViewShapes.Source = value?.Shapes;

                collectionViewShapes.SortDescriptions.Clear();
                collectionViewShapes.SortDescriptions.Add(new SortDescription(nameof(ShapeModel.Icon), ListSortDirection.Ascending));
                collectionViewShapes.SortDescriptions.Add(new SortDescription(nameof(ShapeModel.Description), ListSortDirection.Ascending));
                OnPropertyChanged(nameof(CollectionViewShapes));
            }
        }
        #endregion

        #region SelectedShape : Выбранный shape объект
        /// <summary> Выбранный shape объект </summary>
        private ShapeModel selectedShape;
        /// <summary> Выбранный shape объект </summary>
        public ShapeModel SelectedShape
        {
            get => selectedShape;
            set
            {
                if (!Set(ref selectedShape, value)) return;
                if (selectedShape == null) return;
                if (corelApp.Documents.Count == 0) return;
                SelectShape();
            }
        }
        #endregion

        #region ElapsedTime : Затраченное время
        /// <summary> Затраченное время </summary>
        private string elapsedTime;
        /// <summary> Затраченное время </summary>
        public string ElapsedTime
        {
            get => elapsedTime;
            set => Set(ref elapsedTime, value);
        }
        #endregion

        #endregion

        #region Команды

        #region ScanAllShapes : Cканировать все shape объекты
        /// <summary> Cканировать все shape объекты </summary>
        public ICommand ScanAllShapesCommand { get; }
        private bool CanScanAllShapesCommandExecute(object p) => true;
        private void OnScanAllShapesCommandExecuted(object p)
        {
            if (corelApp.Documents.Count == 0)
            {
                ClearCollection();
                return;
            }

            stopwatch.Restart();

            corelApp.BeginDraw(commandGroup: false);

            FiltersManger filtersManger = new FiltersManger(corelApp.ActiveDocument.Index);
            CorelDrawHandler corelDrawHandler = new CorelDrawHandler(corelApp, filtersManger);

            foreach (corel.Page page in corelApp.ActiveDocument.Pages)
            {
                corelDrawHandler.Scan(page.Shapes.All(), page.Index);
            }

            filters = filtersManger.GetFilters();
            collectionViewFilters.Source = filters;

            collectionViewFilters.GroupDescriptions.Clear();
            collectionViewFilters.GroupDescriptions.Add(new PropertyGroupDescription(nameof(FilterModel.GroupName)));
            OnPropertyChanged(nameof(CollectionViewFilters));

            stopwatch.Stop();
            ElapsedTime = stopwatch.ElapsedMilliseconds.ToString() + " ms";

            corelApp.EndDraw();
        }
        #endregion

        #endregion

        #region ctor
        public DockerUIViewModel(corel.Application app)
        {
            corelApp = app;

            #region Команды

            ScanAllShapesCommand = new LambdaCommand(OnScanAllShapesCommandExecuted, CanScanAllShapesCommandExecute);

            #endregion

            #region Тестовые данные

            List<FilterModel> filters = new List<FilterModel>();

            filters.Add(new FilterModel()
            {
                Icon = InfIconType.CMYKColorModel,
                Description = "CMYK Color",
                Shapes = Enumerable.Range(1, 200).Select(i => new ShapeModel(i, 1, 1)
                {
                    Description = "C:100, M:100, Y:100, K:100",
                    Icon = InfIconType.CMYKColorModel,
                    SecondIconVisibility = true,
                    ColorHexValue = "#00FF00"
                }).ToList()
            });

            filters.Add(new FilterModel()
            {
                Icon = InfIconType.RGBColorModel,
                Description = "RGB Color",
                Shapes = Enumerable.Range(1, 200).Select(i => new ShapeModel(i, 1, 1)
                {
                    Description = "R:100, G:100, B:100",
                    Icon = InfIconType.RGBColorModel,
                    SecondIconVisibility = true,
                    ColorHexValue = "#00FF00"
                }).ToList()
            });

            collectionViewFilters.Source = filters;
            OnPropertyChanged(nameof(CollectionViewFilters));

            #endregion

            stopwatch = new Stopwatch();

            corelApp.DocumentClose += CorelApp_DocumentClose;
        }
        #endregion

        #region События corelApp

        private void CorelApp_DocumentClose(corel.Document Doc)
        {
            ClearCollection();
        }

        #endregion

        #region Services

        private void SelectShape()
        {
            corel.ShapeRange sr = corelApp.Documents[selectedShape.DocID].Pages[selectedShape.PageID].Shapes.All();
            corel.Shape shape = null;
            FindShape(selectedShape.ID, sr, ref shape);

            if (shape != null)
            {
                shape.Application.Documents[selectedShape.DocID].Activate();
                shape.Page.Activate();
                shape.CreateSelection();
            }
            else
            {
                selectedFilter.Shapes.Remove(selectedShape);
                if (selectedFilter.Shapes.Count == 0)
                    filters.Remove(SelectedFilter);
                CollectionViewShapes.Refresh();
                CollectionViewFilters.Refresh();
            }
        }

        private void ClearCollection()
        {
            filters = null;
            selectedFilter = null;
            collectionViewFilters.Source = null;
            collectionViewShapes.Source = null;
            OnPropertyChanged(nameof(CollectionViewFilters));
            OnPropertyChanged(nameof(collectionViewShapes));
        }

        private void FindShape(int shapeID, corel.ShapeRange sr, ref corel.Shape shape)
        {
            bool done = false;

            foreach (corel.Shape s in sr)
            {
                if (done) return;

                if (s.StaticID == shapeID)
                {
                    shape = s;
                    done = true;
                }

                if (s.Type == corel.cdrShapeType.cdrGroupShape)
                    FindShape(shapeID, s.Shapes.All(), ref shape);

                if (s.PowerClip != null)
                    FindShape(shapeID, s.PowerClip.Shapes.All(), ref shape);
            }
        }

        #endregion
    }
}
