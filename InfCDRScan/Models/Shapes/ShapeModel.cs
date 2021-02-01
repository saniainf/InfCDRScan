using InfCDRScan.Services;
using corel = Corel.Interop.VGCore;

namespace InfCDRScan.Models.Shapes
{
    internal class ShapeModel
    {
        public readonly int ID;
        public readonly int PageID;
        public readonly int DocID;
        public string Description { get; set; }
        public bool FirstIconVisibility { get => Icon != InfIconType.None; }
        public InfIconType Icon { get; set; }
        public bool SecondIconVisibility { get; set; }
        public string ColorHexValue { get; set; }

        public ShapeModel(int shapeID, int pageID, int docID)
        {
            ID = shapeID;
            PageID = pageID;
            DocID = docID;
        }
    }
}
