using System;
using System.Collections.Generic;
using corel = Corel.Interop.VGCore;
using InfCDRScan.Models.Shapes;
using InfCDRScan.Services;

namespace InfCDRScan.Models.Filters
{
    internal class FilterModel
    {
        public InfIconType Icon { get; set; }
        public string Description { get; set; }
        public string GroupName { get; set; }
        public ICollection<ShapeModel> Shapes { get; set; }
        public string Count => Shapes.Count.ToString();
    }
}
