﻿using System;
using System.ComponentModel.DataAnnotations;

namespace OneDriveForBusiness.MVC.ADAL.Model
{
    public class UserTokenCache
    {
        [Key]
        public int UserTokenCacheId { get; set; }
        public string webUserUniqueId { get; set; }
        public byte[] cacheBits { get; set; }
        public DateTime LastWrite { get; set; }
    }
}