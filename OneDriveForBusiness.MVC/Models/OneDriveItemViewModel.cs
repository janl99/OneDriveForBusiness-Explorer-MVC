using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OneDriveForBusiness.MVC.Models
{
    public class OneDriveItemViewModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public long Size { get; set; }
        public string Creator { get; set; }
        public string Extension { get; set; }
    }
}