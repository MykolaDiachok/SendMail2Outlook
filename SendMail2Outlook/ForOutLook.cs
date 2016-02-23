using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SendMailOutlook
{
    [JsonObject]
    class ForOutLook
    {
        [JsonProperty("subject")]
        public string subject { get; set; }
        [JsonProperty("body")]
        public string body { get; set; }
        [JsonProperty("TO")]
        public List<string> TO { get; set; }
        [JsonProperty("Attachments")]
        public List<ForOutLookAttachments> Attachments { get; set; }
    }

    class ForOutLookTo
    {
        public string To { get; set; }
    }

    class ForOutLookAttachments
    {
        public string filename { get; set; }
        public string Base64 { get; set; }
    }
}
