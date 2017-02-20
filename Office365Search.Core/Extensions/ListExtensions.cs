using Office365Search.Core.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365Search.Core.Extensions
{
    public static class ListExtensions
    {
        //public async static Task<bool> CheckInAsync(this DocumentInformation document)
        //{
        //    bool successful = false;

        //    string siteUrl = document.SiteUrl.ToLower().EnsureEndsWith("/");
        //    Guid itemId = document.FileId;
        //    string checkInComment = "Checked in by Cortana";

        //    var sharePointCredentials = Helpers.SettingsHelper.GetSharePointCredentials();

        //    string userName = sharePointCredentials.UserName;
        //    string password = sharePointCredentials.Password;

        //    try
        //    {
        //        successful = await Helpers.ContextHelper.CheckInDocumentAsync(siteUrl, userName, password, itemId, checkInComment);
        //    }
        //    catch (Exception ex)
        //    {
        //    }

        //    return successful;
        //}


        //public static List<string> AsPasswordList(this List<string> terms)
        //{
        //    List<string> passwords = new List<string>();

        //    passwords = (from term in terms
        //                 select term.AsPassword().AsTitleCase()).ToList();

        //    return passwords;
        //}

        //public static List<LocationResultInformation> ToSearchResultList(this List<Json.RawLocationResultInformation> list)
        //{
        //    if (list == null) list = new List<Json.RawLocationResultInformation>();

        //    return (from item in list
        //            select item.ToSearchResult()).OrderBy(o => o.DistanceAway).ToList();
        //}

        //public static List<Json.RawLocationResultInformation> ToSearchResultList(this string data)
        //{
        //    MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(data));

        //    DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(List<Json.RawLocationResultInformation>));
        //    var result = ser.ReadObject(ms) as List<Json.RawLocationResultInformation>;

        //    return result;

        //}

        public static ContentClassInformation AsContentClassInformation(this string contentClass)
        {
            if (string.IsNullOrEmpty(contentClass))
            {
                return null;
            }
            else
            {
                List<string> contentClassInformation = contentClass.Split("_".ToCharArray()).ToList();

                return new ContentClassInformation()
                {
                    Root = contentClassInformation.FirstOrDefault(),
                    Type = (contentClassInformation.Count > 0) ? contentClassInformation[1] : null,
                    SubType = (contentClassInformation.Count > 1) ? contentClassInformation[2] : null,
                };
            }
        }


        //public static LocationResultInformation ToSearchResult(this Json.RawLocationResultInformation result)
        //{
        //    return new LocationResultInformation()
        //    {
        //        Id = result.Id,
        //        DisplayName = result.DisplayName,
        //        DistanceAway = result.DistanceAway,
        //        District = result.District,
        //        FormattedAddress = result.FormattedAddress,
        //        HoursOfOperation = result.HoursOfOperation,
        //        ImageUrl = result.ImageUrl,
        //        Locality = result.Locality,
        //        PhoneNumber = result.PhoneNumber,
        //        PostalCode = result.PostalCode,
        //        Region = result.Region,
        //        StreetAddress = result.StreetAddress,
        //        WebsiteUrl = result.WebsiteUrl,
        //        Latitude = result.Latitude,
        //        Longitude = result.Longitude,
        //    };
        //}

        //public static string DeriveStreetAddress(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.BusinessAddress != null)
        //    {
        //        return contact.BusinessAddress.Street + "";
        //    }
        //    else if (contact.HomeAddress != null)
        //    {
        //        return contact.HomeAddress.Street + "";
        //    }
        //    else if (contact.OtherAddress != null)
        //    {
        //        return contact.OtherAddress.Street + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string DeriveCity(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.BusinessAddress != null)
        //    {
        //        return contact.BusinessAddress.City + "";
        //    }
        //    else if (contact.HomeAddress != null)
        //    {
        //        return contact.HomeAddress.City + "";
        //    }
        //    else if (contact.OtherAddress != null)
        //    {
        //        return contact.OtherAddress.City + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string DeriveState(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.BusinessAddress != null)
        //    {
        //        return contact.BusinessAddress.State + "";
        //    }
        //    else if (contact.HomeAddress != null)
        //    {
        //        return contact.HomeAddress.State + "";
        //    }
        //    else if (contact.OtherAddress != null)
        //    {
        //        return contact.OtherAddress.State + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string DeriveCountry(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.BusinessAddress != null)
        //    {
        //        return contact.BusinessAddress.CountryOrRegion + "";
        //    }
        //    else if (contact.HomeAddress != null)
        //    {
        //        return contact.HomeAddress.CountryOrRegion + "";
        //    }
        //    else if (contact.OtherAddress != null)
        //    {
        //        return contact.OtherAddress.CountryOrRegion + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string DerivePostalCode(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.BusinessAddress != null)
        //    {
        //        return contact.BusinessAddress.PostalCode + "";
        //    }
        //    else if (contact.HomeAddress != null)
        //    {
        //        return contact.HomeAddress.PostalCode + "";
        //    }
        //    else if (contact.OtherAddress != null)
        //    {
        //        return contact.OtherAddress.PostalCode + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string DeriveEmailAddress(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.EmailAddresses != null)
        //    {
        //        return contact.EmailAddresses.FirstOrDefault() + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}

        //public static string DerivePhoneNumber(this OfficePoint.Core.Json.Contacts.Value contact)
        //{
        //    if (contact.BusinessPhones != null && contact.BusinessPhones.Count() > 0)
        //    {
        //        return contact.BusinessPhones.FirstOrDefault() + "";
        //    }
        //    else if (contact.HomePhones != null && contact.HomePhones.Count() > 0)
        //    {
        //        return contact.HomePhones.FirstOrDefault() + "";
        //    }
        //    else
        //    {
        //        return "";
        //    }
        //}
    }

}
