using System;

namespace Xbim.CobieExpress.Exchanger.FilterHelper
{
    public static class RoleFilterExtensions
    {
        public static string ToResourceName(this RoleFilter filter)
        {
            const string format = "Xbim.CobieExpress.Exchanger.FilterHelper.COBie{0}Filters.config";
            return filter == RoleFilter.Unknown 
                ? string.Format(format, "Default") 
                : string.Format(format, filter);
        }

        public static bool HasMultipleFlags(this RoleFilter filter)
        {
            int flagsSet = 0;
            foreach (RoleFilter flag in Enum.GetValues(typeof(RoleFilter)))
            {
                if(filter.HasFlag(flag))
                {
                    flagsSet++;
                    if(flagsSet > 1)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}