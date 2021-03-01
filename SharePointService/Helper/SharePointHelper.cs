using Interfaces;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Reflection;
using static Helper.ReflactionHelper;

namespace Helper
{
    public static class SharePointHelper
    {
        public static ListItem CreateListItem<T>(T item, List selectedList)
            where T : IBaseModel
        {
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem newListItem = selectedList.AddItem(itemCreateInfo);

            PropertyInfo[] props = typeof(T).GetProperties();
            foreach (var prop in props)
            {
                if (prop.GetValue(item) is object value)
                    newListItem[prop.Name] = value;
            }
            return newListItem;
        }

        public static void UpdateListItem<T>(T item, ListItem selectedItem) 
            where T : IBaseModel
        {
            PropertyInfo[] props = typeof(T).GetProperties();
            foreach (var prop in props)
            {
                if (prop.Name == "ID")
                    continue;

                if (prop.GetValue(item) is object value)
                    selectedItem[prop.Name] = value;
            }
        }

        public static List<T> ParseSharePointList<T>(ListItemCollection collListItem) 
            where T : IBaseModel, new ()
        {
            var output = new List<T>();
            foreach (var item in collListItem)
            {
                var _tObject = new T();
                PropertyInfo[] props = typeof(T).GetProperties();
                foreach (var prop in props)
                {
                    var value = item.FieldValues[prop.Name];
                    if (value is null)
                        continue;

                    Type propertyType = GetPropertyType(prop);

                    prop.SetValue(_tObject, Convert.ChangeType(value, propertyType));
                }
                output.Add(_tObject);
            }

            return output;
        }



    }
}
