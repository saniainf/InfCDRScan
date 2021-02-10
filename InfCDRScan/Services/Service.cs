using System;

namespace InfCDRScan.Services
{
    internal class Service
    {

    }

    internal class DescriptionAttribute : Attribute
    {
        public string Value { get; private set; }
        public DescriptionAttribute(string value)
        {
            Value = value;
        }
    }

    internal class GroupNameAttribute : Attribute
    {
        public InfFilterGroup Value { get; private set; }
        public GroupNameAttribute(InfFilterGroup value)
        {
            Value = value;
        }
    }

    internal class IconAttribute : Attribute
    {
        public InfIconType Value { get; private set; }
        public IconAttribute(InfIconType value)
        {
            Value = value;
        }
    }

    internal class SortPropAttribute : Attribute
    {
        public string Value { get; private set; }
        public SortPropAttribute(string value)
        {
            Value = value;
        }
    }

    internal class ShapeDataSheet
    {
        public readonly int PageID;
        public readonly int ShapeID;
        public InfFilters FiltersType { get; set; }
        public string Description { get; set; }
        public InfIconType Icon { get; set; }
        public string ColorHexValue { get; set; }

        public ShapeDataSheet(int shapeID, int pageID)
        {
            ShapeID = shapeID;
            PageID = pageID;
        }
    }
}
