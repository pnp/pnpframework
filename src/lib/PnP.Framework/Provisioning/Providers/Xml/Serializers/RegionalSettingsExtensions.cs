using System;
using System.Collections.Generic;
using System.Text;

namespace PnP.Framework.Provisioning.Providers.Xml
{
    internal static class RegionalSettingsExtensions
    {
        public static V202103.CalendarType FromTemplateToSchemaCalendarTypeV201605(this Microsoft.SharePoint.Client.CalendarType calendarType)
        {
            switch (calendarType)
            {
                case Microsoft.SharePoint.Client.CalendarType.ChineseLunar:
                    return V202103.CalendarType.ChineseLunar;
                case Microsoft.SharePoint.Client.CalendarType.Gregorian:
                    return V202103.CalendarType.Gregorian;
                case Microsoft.SharePoint.Client.CalendarType.GregorianArabic:
                    return V202103.CalendarType.GregorianArabicCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianMEFrench:
                    return V202103.CalendarType.GregorianMiddleEastFrenchCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianXLITEnglish:
                    return V202103.CalendarType.GregorianTransliteratedEnglishCalendar;
                case Microsoft.SharePoint.Client.CalendarType.GregorianXLITFrench:
                    return V202103.CalendarType.GregorianTransliteratedFrenchCalendar;
                case Microsoft.SharePoint.Client.CalendarType.Hebrew:
                    return V202103.CalendarType.Hebrew;
                case Microsoft.SharePoint.Client.CalendarType.Hijri:
                    return V202103.CalendarType.Hijri;
                case Microsoft.SharePoint.Client.CalendarType.Japan:
                    return V202103.CalendarType.Japan;
                case Microsoft.SharePoint.Client.CalendarType.Korea:
                    return V202103.CalendarType.Korea;
                case Microsoft.SharePoint.Client.CalendarType.KoreaJapanLunar:
                    return V202103.CalendarType.KoreaandJapaneseLunar;
                case Microsoft.SharePoint.Client.CalendarType.SakaEra:
                    return V202103.CalendarType.SakaEra;
                case Microsoft.SharePoint.Client.CalendarType.Taiwan:
                    return V202103.CalendarType.Taiwan;
                case Microsoft.SharePoint.Client.CalendarType.Thai:
                    return V202103.CalendarType.Thai;
                case Microsoft.SharePoint.Client.CalendarType.UmAlQura:
                    return V202103.CalendarType.UmmalQura;
                case Microsoft.SharePoint.Client.CalendarType.None:
                default:
                    return V202103.CalendarType.None;
            }
        }

        public static Microsoft.SharePoint.Client.CalendarType FromSchemaToTemplateCalendarTypeV201605(this V202103.CalendarType calendarType)
        {
            switch (calendarType)
            {
                case V202103.CalendarType.ChineseLunar:
                    return Microsoft.SharePoint.Client.CalendarType.ChineseLunar;
                case V202103.CalendarType.Gregorian:
                    return Microsoft.SharePoint.Client.CalendarType.Gregorian;
                case V202103.CalendarType.GregorianArabicCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianArabic;
                case V202103.CalendarType.GregorianMiddleEastFrenchCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianMEFrench;
                case V202103.CalendarType.GregorianTransliteratedEnglishCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianXLITEnglish;
                case V202103.CalendarType.GregorianTransliteratedFrenchCalendar:
                    return Microsoft.SharePoint.Client.CalendarType.GregorianXLITFrench;
                case V202103.CalendarType.Hebrew:
                    return Microsoft.SharePoint.Client.CalendarType.Hebrew;
                case V202103.CalendarType.Hijri:
                    return Microsoft.SharePoint.Client.CalendarType.Hijri;
                case V202103.CalendarType.Japan:
                    return Microsoft.SharePoint.Client.CalendarType.Japan;
                case V202103.CalendarType.Korea:
                    return Microsoft.SharePoint.Client.CalendarType.Korea;
                case V202103.CalendarType.KoreaandJapaneseLunar:
                    return Microsoft.SharePoint.Client.CalendarType.KoreaJapanLunar;
                case V202103.CalendarType.SakaEra:
                    return Microsoft.SharePoint.Client.CalendarType.SakaEra;
                case V202103.CalendarType.Taiwan:
                    return Microsoft.SharePoint.Client.CalendarType.Taiwan;
                case V202103.CalendarType.Thai:
                    return Microsoft.SharePoint.Client.CalendarType.Thai;
                case V202103.CalendarType.UmmalQura:
                    return Microsoft.SharePoint.Client.CalendarType.UmAlQura;
                case V202103.CalendarType.None:
                default:
                    return Microsoft.SharePoint.Client.CalendarType.None;
            }
        }

        public static V202103.WorkHour FromTemplateToSchemaWorkHourV201605(this Model.WorkHour workHour)
        {
            switch (workHour)
            {
                case Model.WorkHour.AM0100:
                    return V202103.WorkHour.Item100AM;
                case Model.WorkHour.AM0200:
                    return V202103.WorkHour.Item200AM;
                case Model.WorkHour.AM0300:
                    return V202103.WorkHour.Item300AM;
                case Model.WorkHour.AM0400:
                    return V202103.WorkHour.Item400AM;
                case Model.WorkHour.AM0500:
                    return V202103.WorkHour.Item500AM;
                case Model.WorkHour.AM0600:
                    return V202103.WorkHour.Item600AM;
                case Model.WorkHour.AM0700:
                    return V202103.WorkHour.Item700AM;
                case Model.WorkHour.AM0800:
                    return V202103.WorkHour.Item800AM;
                case Model.WorkHour.AM0900:
                    return V202103.WorkHour.Item900AM;
                case Model.WorkHour.AM1000:
                    return V202103.WorkHour.Item1000AM;
                case Model.WorkHour.AM1100:
                    return V202103.WorkHour.Item1100AM;
                case Model.WorkHour.AM1200:
                    return V202103.WorkHour.Item1200AM;
                case Model.WorkHour.PM0100:
                    return V202103.WorkHour.Item100PM;
                case Model.WorkHour.PM0200:
                    return V202103.WorkHour.Item200PM;
                case Model.WorkHour.PM0300:
                    return V202103.WorkHour.Item300PM;
                case Model.WorkHour.PM0400:
                    return V202103.WorkHour.Item400PM;
                case Model.WorkHour.PM0500:
                    return V202103.WorkHour.Item500PM;
                case Model.WorkHour.PM0600:
                    return V202103.WorkHour.Item600PM;
                case Model.WorkHour.PM0700:
                    return V202103.WorkHour.Item700PM;
                case Model.WorkHour.PM0800:
                    return V202103.WorkHour.Item800PM;
                case Model.WorkHour.PM0900:
                    return V202103.WorkHour.Item900PM;
                case Model.WorkHour.PM1000:
                    return V202103.WorkHour.Item1000PM;
                case Model.WorkHour.PM1100:
                    return V202103.WorkHour.Item1100PM;
                case Model.WorkHour.PM1200:
                    return V202103.WorkHour.Item1200PM;
                default:
                    return V202103.WorkHour.Item100AM;
            }
        }

        public static Model.WorkHour FromSchemaToTemplateWorkHourV201605(this V202103.WorkHour workHour)
        {
            switch (workHour)
            {
                case V202103.WorkHour.Item100AM:
                    return Model.WorkHour.AM0100;
                case V202103.WorkHour.Item200AM:
                    return Model.WorkHour.AM0200;
                case V202103.WorkHour.Item300AM:
                    return Model.WorkHour.AM0300;
                case V202103.WorkHour.Item400AM:
                    return Model.WorkHour.AM0400;
                case V202103.WorkHour.Item500AM:
                    return Model.WorkHour.AM0500;
                case V202103.WorkHour.Item600AM:
                    return Model.WorkHour.AM0600;
                case V202103.WorkHour.Item700AM:
                    return Model.WorkHour.AM0700;
                case V202103.WorkHour.Item800AM:
                    return Model.WorkHour.AM0800;
                case V202103.WorkHour.Item900AM:
                    return Model.WorkHour.AM0900;
                case V202103.WorkHour.Item1000AM:
                    return Model.WorkHour.AM1000;
                case V202103.WorkHour.Item1100AM:
                    return Model.WorkHour.AM1100;
                case V202103.WorkHour.Item1200AM:
                    return Model.WorkHour.AM1200;
                case V202103.WorkHour.Item100PM:
                    return Model.WorkHour.PM0100;
                case V202103.WorkHour.Item200PM:
                    return Model.WorkHour.PM0200;
                case V202103.WorkHour.Item300PM:
                    return Model.WorkHour.PM0300;
                case V202103.WorkHour.Item400PM:
                    return Model.WorkHour.PM0400;
                case V202103.WorkHour.Item500PM:
                    return Model.WorkHour.PM0500;
                case V202103.WorkHour.Item600PM:
                    return Model.WorkHour.PM0600;
                case V202103.WorkHour.Item700PM:
                    return Model.WorkHour.PM0700;
                case V202103.WorkHour.Item800PM:
                    return Model.WorkHour.PM0800;
                case V202103.WorkHour.Item900PM:
                    return Model.WorkHour.PM0900;
                case V202103.WorkHour.Item1000PM:
                    return Model.WorkHour.PM1000;
                case V202103.WorkHour.Item1100PM:
                    return Model.WorkHour.PM1100;
                case V202103.WorkHour.Item1200PM:
                    return Model.WorkHour.PM1200;
                default:
                    return Model.WorkHour.AM0100;
            }
        }
    }
}
