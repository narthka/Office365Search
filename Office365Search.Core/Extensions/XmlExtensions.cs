using Office365Search.Core.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Office365Search.Core.Extensions
{
    public static class XmlExtensions
    {
        public static ManagerInformation ParseToManager(this XDocument document)
        {
            long managerId = 0;

            XNamespace atom = "http://www.w3.org/2005/Atom";
            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";
            XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";

            List<XElement> items = document.Descendants(d + "query").Elements(d + "PrimaryQueryResult")
                                                   .Elements(d + "RelevantResults")
                                                   .Elements(d + "Table")
                                                   .Elements(d + "Rows")
                                                   .Elements(d + "element")
                                                   .ToList();

            var managerTitleValue = ((items.Descendants(d + "Key").ToList().Where(w => w.Value == "Title").First() as XElement).NextNode as XElement).Value;
            var managerUserNameValue = ((items.Descendants(d + "Key").ToList().Where(w => w.Value == "UserName").First() as XElement).NextNode as XElement).Value;
            var managerIdValue = ((items.Descendants(d + "Key").ToList().Where(w => w.Value == "DocId").First() as XElement).NextNode as XElement).Value;

            long.TryParse(managerIdValue + "", out managerId);

            return new ManagerInformation()
            {
                DisplayName = managerTitleValue,
                EmailAddress = managerUserNameValue,
                Id = managerId,
            };
        }

        private static List<string> ToStringList(this XElement document, string propertyName, string delimiter)
        {
            List<string> list = new List<string>();

            var node = document.Elements().Elements().FirstOrDefault(e => e.Value.Equals(propertyName)) as XElement;

            var nodeValue = (node == null || string.IsNullOrEmpty((node.NextNode as XElement).Value)) ? "" : (node.NextNode as XElement).Value;

            return nodeValue.Split(delimiter.ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList();
        }
        private static long ToLongValue(this XElement document, string propertyName)
        {
            var node = document.Elements().Elements().FirstOrDefault(e => e.Value.Equals(propertyName)) as XElement;

            return (node == null || string.IsNullOrEmpty((node.NextNode as XElement).Value)) ? 0 : Convert.ToInt64((node.NextNode as XElement).Value);
        }

        private static string ToStringValue(this XElement document, string propertyName)
        {
            var node = document.Elements().Elements().FirstOrDefault(e => e.Value.Equals(propertyName)) as XElement;

            return (node == null) ? null : (node.NextNode as XElement).Value;
        }

        private static Guid ToGuidValue(this XElement document, string propertyName)
        {
            var node = document.Elements().Elements().FirstOrDefault(e => e.Value.Equals(propertyName)) as XElement;

            string nodeValue = (node == null) ? null : (node.NextNode as XElement).Value;

            return (string.IsNullOrEmpty(nodeValue)) ? Guid.Empty : new Guid(nodeValue);
        }

        private static DateTime ToDateValue(this XElement document, string propertyName)
        {
            var node = document.Elements().Elements().FirstOrDefault(e => e.Value.Equals(propertyName)) as XElement;

            string nodeValue = (node == null) ? null : (node.NextNode as XElement).Value;

            return (string.IsNullOrEmpty(nodeValue)) ? DateTime.MinValue : DateTime.Parse(nodeValue);
        }


        public static bool HasPreview(this XElement element)
        {
            string documentPreviewMetadata = element.ToStringValue("DocumentPreviewMetadata");

            return (!string.IsNullOrEmpty(documentPreviewMetadata) && !documentPreviewMetadata.Contains("-1x-1x0"));
        }

        public static AuthorInformation AsAuthorInformation(this string author)
        {
            if (string.IsNullOrEmpty(author))
            {
                return null;
            }
            else
            {
                List<string> authorInformation = author.Split("|".ToCharArray()).ToList();

                return new AuthorInformation()
                {
                    DisplayName = authorInformation[1].Trim(),
                    EmailAddress = authorInformation[0].Trim(),
                    LoginName = ParseLoginName(authorInformation),
                    Id = authorInformation[2].Trim().Split(" ".ToCharArray()).FirstOrDefault(),
                };
            }
        }

        private static string ParseLoginName(List<string> authorInformation)
        {
            return "";
        }

        public static string ToFileIconUrl(this string fileExtension)
        {
            string imageUrl = "";

            string extension = fileExtension.ToLower().Substring(0, Math.Min(fileExtension.Length, 3));

            switch (extension)
            {
                case "xls":
                    imageUrl = "/images/fileicons/256_ICXLSX.PNG";
                    break;
                case "doc":
                    imageUrl = "/images/fileicons/256_ICDOCX.PNG";
                    break;
                case "ppt":
                    imageUrl = "/images/fileicons/256_ICPPTX.PNG";
                    break;
                default:
                    imageUrl = "/images/fileicons/256_ICGEN.png";
                    break;
            }

            return imageUrl;
        }

        public static List<DocumentInformation> ToDocumentList(this XDocument document, bool isGraphQuery)
        {
            List<DocumentInformation> documents = new List<DocumentInformation>();

            XNamespace dns = "http://schemas.microsoft.com/ado/2007/08/dataservices";

            documents = (from doc in document.Descendants(dns + "Cells")
                         select new DocumentInformation()
                         {
                             Title = doc.ToStringValue("Title"),
                             AuthorDisplayNames = doc.ToStringValue("Author").Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList(),
                             AuthorInformation = doc.ToStringValue("AuthorOwsUser").AsAuthorInformation(),
                             DocumentPreviewInformation = doc.ToStringValue("DocumentPreviewMetadata"),
                             HasPreview = doc.HasPreview(),
                             DocumentId = doc.ToLongValue("DocId"),
                             FileExtension = doc.ToStringValue("FileExtension"),
                             RedirectedURL = doc.ToStringValue("ServerRedirectedURL"),
                             ModifiedBy = doc.ToStringValue("EditorOwsUser").AsAuthorInformation(),
                             ModifiedDate = doc.ToDateValue("LastModifiedTime"),
                             SiteTitle = doc.ToStringValue("SiteTitle"),
                             SiteUrl = doc.ToStringValue("SPWebUrl"),

                             IconUrl = doc.ToStringValue("FileExtension").ToFileIconUrl(),

                             Url = doc.ToStringValue("Path"),
                             SiteId = doc.ToGuidValue("siteID"),
                             WebId = doc.ToGuidValue("webID"),
                             ListId = doc.ToGuidValue("ListID"),
                             FileId = doc.ToGuidValue("uniqueID"),
                             ItemId = doc.ToLongValue("ListItemID"),

                             ThumbnailUrl = doc.ToStringValue("PictureThumbnailURL"),

                             ContentClass = doc.ToStringValue("ContentClass").AsContentClassInformation(),

                             ViewCount = doc.ToLongValue("ViewCountLifetime"),

                         }).Where(w => w.HasPreview).ToList();

            return documents;
        }

        public static List<DocumentInformation> ToDocumentList(this XDocument document)
        {
            List<DocumentInformation> documents = new List<DocumentInformation>();

            XNamespace dns = "http://schemas.microsoft.com/ado/2007/08/dataservices";

            try
            {

                documents = (from doc in document.Descendants(dns + "Cells")
                             select new DocumentInformation()
                             {
                                 Title = doc.ToStringValue("Title"),
                                 AuthorDisplayNames = doc.ToStringValue("Author").Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).ToList(),
                                 AuthorInformation = doc.ToStringValue("AuthorOwsUser").AsAuthorInformation(),
                                 DocumentPreviewInformation = doc.ToStringValue("DocumentPreviewMetadata"),
                                 HasPreview = doc.HasPreview(),
                                 DocumentId = doc.ToLongValue("DocId"),
                                 FileExtension = doc.ToStringValue("FileExtension"),

                                 IconUrl = doc.ToStringValue("FileExtension").ToFileIconUrl(),

                                 RedirectedURL = doc.ToStringValue("ServerRedirectedURL"),
                                 ModifiedBy = doc.ToStringValue("EditorOwsUser").AsAuthorInformation(),
                                 ModifiedDate = doc.ToDateValue("LastModifiedTime"),
                                 SiteTitle = doc.ToStringValue("SiteTitle"),//.CheckAsOneDrive(doc.ToStringValue("ContentClass")),
                                 SiteUrl = doc.ToStringValue("SPWebUrl"),

                                 Url = doc.ToStringValue("Path"),
                                 SiteId = doc.ToGuidValue("siteID"),
                                 WebId = doc.ToGuidValue("webID"),
                                 ListId = doc.ToGuidValue("ListID"),
                                 FileId = doc.ToGuidValue("uniqueID"),
                                 ItemId = doc.ToLongValue("ListItemID"),

                                 CheckedOutTo = doc.ToStringList("CheckoutUserOWSUSER", "|"),

                                 ThumbnailUrl = doc.ToStringValue("PictureThumbnailURL"),

                                 IsDocument = doc.ToStringValue("IsDocument").Equals("TRUE", StringComparison.OrdinalIgnoreCase),
                                 ContentClass = doc.ToStringValue("ContentClass").AsContentClassInformation(),

                                 ViewCount = doc.ToLongValue("ViewCountLifetime"),

                             }).ToList();
            }
            catch (Exception ex)
            {

            }

            return documents;
        }

    }

}
