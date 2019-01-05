using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Hiz.Npoi.Attributes
{
    // 代码来至: System.ComponentModel.DataAnnotations.LocalizableString
    class LocalizableString
    {
        string _PropertyName;
        public LocalizableString(string propertyName)
        {
            this._PropertyName = propertyName;
        }

        string _PropertyValue;
        public string Value
        {
            get
            {
                return this._PropertyValue;
            }
            set
            {
                if (this._PropertyValue != value)
                {
                    this.ClearCache();
                    this._PropertyValue = value;
                }
            }
        }

        Type _ResourceType;
        public Type ResourceType
        {
            get
            {
                return this._ResourceType;
            }
            set
            {
                if (this._ResourceType != value)
                {
                    this.ClearCache();
                    this._ResourceType = value;
                }
            }
        }

        Func<string> _CachedResult;
        void ClearCache()
        {
            this._CachedResult = null;
        }

        const string LocalizationFailed‎ = "Cannot retrieve property '{0}' because localization failed.  Type '{1}' is not public or does not contain a public static string property with the name '{2}'.";
        public string GetLocalizableValue()
        {
            if (this._CachedResult == null)
            {
                if (this._PropertyValue == null || this._ResourceType == null)
                {
                    this._CachedResult = () => this._PropertyValue;
                }
                else
                {
                    var property = this._ResourceType.GetProperty(this._PropertyValue);
                    var flag = false;
                    if (!this._ResourceType.IsVisible || property == null || property.PropertyType != typeof(string))
                    {
                        flag = true;
                    }
                    else
                    {
                        var getter = property.GetGetMethod();
                        if (getter == null || !getter.IsPublic || !getter.IsStatic)
                        {
                            flag = true;
                        }
                    }
                    if (flag)
                    {
                        var message = string.Format(CultureInfo.CurrentCulture, LocalizationFailed, new object[]
                        {
                            this._PropertyName,
                            this._ResourceType.FullName,
                            this._PropertyValue
                        });
                        this._CachedResult = () => throw new InvalidOperationException(message);
                    }
                    else
                    {
                        this._CachedResult = () => (string)property.GetValue(null, null);
                    }
                }
            }
            return this._CachedResult();
        }
    }
}
