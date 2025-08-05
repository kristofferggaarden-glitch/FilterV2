using System;
using System.Collections.Generic;

namespace FilterV1
{
    public class CellPairComparer : IEqualityComparer<RemoveCellsWindow.CellPair>, IEqualityComparer<AddTextWindow.CellPair>
    {
        public bool Equals(RemoveCellsWindow.CellPair x, RemoveCellsWindow.CellPair y)
        {
            if (x == null || y == null)
                return false;
            return string.Equals(x.FirstCell, y.FirstCell, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(x.SecondCell, y.SecondCell, StringComparison.OrdinalIgnoreCase);
        }

        public bool Equals(AddTextWindow.CellPair x, AddTextWindow.CellPair y)
        {
            if (x == null || y == null)
                return false;
            return string.Equals(x.FirstCell, y.FirstCell, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(x.SecondCell, y.SecondCell, StringComparison.OrdinalIgnoreCase);
        }

        public int GetHashCode(RemoveCellsWindow.CellPair obj)
        {
            if (obj == null) return 0;
            return StringComparer.OrdinalIgnoreCase.GetHashCode(obj.FirstCell ?? "") ^
                   StringComparer.OrdinalIgnoreCase.GetHashCode(obj.SecondCell ?? "");
        }

        public int GetHashCode(AddTextWindow.CellPair obj)
        {
            if (obj == null) return 0;
            return StringComparer.OrdinalIgnoreCase.GetHashCode(obj.FirstCell ?? "") ^
                   StringComparer.OrdinalIgnoreCase.GetHashCode(obj.SecondCell ?? "");
        }
    }
}