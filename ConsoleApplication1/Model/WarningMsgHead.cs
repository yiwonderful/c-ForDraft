using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Runtime.Serialization;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace BS.Alarm.WebApi.Models
{
    [DataContract]
    [DisplayName("预警主表")]
    public class WarningMsgHead
    {
        [Display(Name = "主键ID")]
        [DataMember]
        public string ID { get; set; }

        [Display(Name = "发送对象")]
        [DataMember]
        public string Receiver { get; set; }
    }
}