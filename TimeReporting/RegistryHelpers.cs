using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    class RegistryHelpers
    {
        public static void ReadValue<TType>(ref TType result, Microsoft.Win32.RegistryKey key, string keyName,
            Microsoft.Win32.RegistryValueKind valueKind = Microsoft.Win32.RegistryValueKind.String)
        {
            try
            {
                object obj = key.GetValue(keyName);
                if (obj != null)
                {
                result = (TType)Convert.ChangeType(obj, typeof(TType));
                return;
                }
            }
            catch (Exception)
            {}

            try
            {
                key.SetValue(keyName, result, valueKind);
            }
            catch (Exception)
            {}
        }
    }
}
