using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office365Search.Core.Model
{
    public class ContentClassInformation
    {
        public string Root { get; set; }
        public string Type { get; set; }
        public string SubType { get; set; }
    }

    public class AuthorInformation
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string EmailAddress { get; set; }
        public string LoginName { get; set; }
    }

    public class ManagerInformation
    {
        public long Id { get; set; }
        public string DisplayName { get; set; }
        public string EmailAddress { get; set; }
    }

    public class DocumentCollection : List<DocumentInformation>
    {
        public ManagerInformation Manager { get; set; }
    }


    public class DocumentInformation
    {
        public string Title { get; set; }
        public long DocumentId { get; set; }
        public List<string> AuthorDisplayNames { get; set; }
        public AuthorInformation AuthorInformation { get; set; }
        public AuthorInformation ModifiedBy { get; set; }
        public string FileExtension { get; set; }
        public Guid SiteId { get; set; }
        public string SiteTitle { get; set; }
        public Guid WebId { get; set; }
        public Guid ListId { get; set; }
        public Guid FileId { get; set; }
        public long ItemId { get; set; }
        public string SiteUrl { get; set; }
        public string Url { get; set; }
        public string RedirectedURL { get; set; }
        public string DocumentPreviewInformation { get; set; }
        public bool HasPreview { get; set; }
        public DateTime ModifiedDate { get; set; }
        public bool IsDocument { get; set; }
        public string ThumbnailUrl { get; set; }
        public string IconUrl { get; set; }
        public List<string> CheckedOutTo { get; set; }
        public ContentClassInformation ContentClass { get; set; }
        public long ViewCount { get; set; }

    }

}
