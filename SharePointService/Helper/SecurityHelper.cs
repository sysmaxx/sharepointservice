﻿using System;
using System.Security;

namespace Helper
{
    public static class SecurityHelper
    {
        public static SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }

    }
}
