using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.Contacts;

namespace Office365Search.Core.Extensions
{
    public static class StreamExtensions
    {
        public static async Task<Stream> AsStringStream(this Windows.Web.Http.HttpResponseMessage result)
        {
            var content = await result.Content.ReadAsInputStreamAsync();

            return content.AsStreamForRead();
        }

        public static string EnsureStartsWith(this string value, string startsWith)
        {
            return (value.StartsWith(startsWith)) ? value : startsWith + value;
        }

        //public async static Task<List<ContactInformation>> ToContactsList(this Windows.Web.Http.HttpResponseMessage result)
        //{
        //    List<ContactInformation> contacts = new List<ContactInformation>();

        //    var serializer = new DataContractJsonSerializer(typeof(Json.Contacts.RawContacts));

        //    var content = await result.Content.ReadAsStringAsync();
        //    var stream = new MemoryStream(Encoding.UTF8.GetBytes(content));
        //    var FinalData = serializer.ReadObject(stream) as Json.Contacts.RawContacts;

        //    contacts = (from contact in FinalData.value
        //                select new ContactInformation()
        //                {
        //                    ContactId = contact.Id,
        //                    DisplayName = contact.DisplayName,
        //                    MobilePhoneNumber = (contact.MobilePhone1 + ""),
        //                    StreetAddress = contact.DeriveStreetAddress(),
        //                    City = contact.DeriveCity(),
        //                    State = contact.DeriveState(),
        //                    Country = contact.DeriveCountry(),
        //                    PostalCode = contact.DerivePostalCode(),

        //                    PrimaryPhoneNumber = contact.DerivePhoneNumber(),
        //                    WebsiteUrl = contact.DeriveEmailAddress(),

        //                }).Where(w => !string.IsNullOrEmpty(w.StreetAddress)).ToList();

        //    return contacts;
        //}

        //public async static Task<List<MessageInformation>> ToMessagesList(this Windows.Web.Http.HttpResponseMessage result)
        //{
        //    List<MessageInformation> messages = new List<MessageInformation>();

        //    var serializer = new DataContractJsonSerializer(typeof(Json.Mail.RawMessages));

        //    var content = await result.Content.ReadAsStringAsync();
        //    var stream = new MemoryStream(Encoding.UTF8.GetBytes(content));
        //    var FinalData = serializer.ReadObject(stream) as Json.Mail.RawMessages;
        //    messages = (from item in FinalData.value
        //                select new MessageInformation()
        //                {
        //                    ItemId = item.Id,
        //                    Subject = item.Subject,
        //                    SenderDisplayName = item.Sender.EmailAddress.Name,
        //                    BodyPreview = item.BodyPreview.Split('\r').FirstOrDefault().Replace("\r\n", " "),
        //                    CreatedDate = DateTime.Parse(item.DateTimeCreated),

        //                }).ToList();

        //    return messages;
        //}

    }
}
