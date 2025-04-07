using HHG.Common.Runtime;
using HHG.GoogleSheets.Runtime;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using UnityEditor;
using UnityEngine;

namespace HHG.GoogleSheets.Editor
{
    public class SheetImporterTool : EditorWindow
    {
        [MenuItem("| Half Human Games |/Tools/Import Google Sheets")]
        public static void ImportSheets()
        {
            foreach (System.Type type in GetAllScriptableObjectSheetTypes())
            {
                SheetAttribute attr = type.GetCustomAttribute<SheetAttribute>();

                if (attr == null)
                {
                    continue;
                }

                Debug.Log($"Importing sheet for {type.Name}");

                string csv = DownloadSheetCSV(attr.SpreadsheetId, attr.Gid);
                if (string.IsNullOrEmpty(csv))
                {
                    Debug.LogError($"Failed to download sheet for {type.Name}");
                    continue;
                }

                ImportCSVToScriptableObjects(csv, type, attr.Casing);
            }

            AssetDatabase.SaveAssets();
            Debug.Log("✅ All sheets imported!");
        }

        private static string DownloadSheetCSV(string spreadsheetId, string gid)
        {
            try
            {
                string url = $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/export?format=csv";

                if (!string.IsNullOrEmpty(gid))
                {
                    url += $"&gid={System.Uri.EscapeDataString(gid)}";
                }

                using WebClient client = new WebClient();
                return client.DownloadString(url);
            }
            catch (System.Exception ex)
            {
                Debug.LogError($"Error downloading Google Sheet: {ex.Message}");
                return null;
            }
        }

        private static void ImportCSVToScriptableObjects(string csv, System.Type type, Case casing)
        {
            List<List<string>> rows = CsvUtil.Parse(csv);

            if (rows.Count < 2)
            {
                return;
            }

            List<string> headers = rows[0];

            for (int i = 1; i < rows.Count; i++)
            {
                List<string> values = rows[i];
                Dictionary<string, string> row = headers.Zip(values, (k, v) => new KeyValuePair<string, string>(k, v)).ToDictionary(x => x.Key, x => x.Value);

                if (!row.TryGetValue("Name", out string name) || string.IsNullOrEmpty(name))
                {
                    continue;
                }

                ScriptableObject scriptableObject = FindScriptableObjectByName(type, name);

                if (scriptableObject == null)
                {
                    Debug.LogWarning($"No ScriptableObject found for {type.Name} named '{name}'");
                    continue;
                }

                ApplyFieldsRecursive(scriptableObject, row, type, scriptableObject, casing);
                EditorUtility.SetDirty(scriptableObject);
            }
        }

        private static void ApplyFieldsRecursive(object instance, Dictionary<string, string> row, System.Type type, object rootContext, Case casing)
        {
            foreach (FieldInfo field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                if (field.FieldType.IsAssignableFrom(typeof(Object)))
                {
                    continue;
                }

                SheetFieldAttribute attr = field.GetCustomAttribute<SheetFieldAttribute>();

                if (attr != null)
                {
                    string columnName = attr.ColumnName;

                    if (string.IsNullOrEmpty(columnName))
                    {
                        columnName = casing.ToCase(field.Name);
                    }

                    if (row.TryGetValue(columnName, out string raw))
                    {
                        object value = ConvertValue(columnName, raw, field.FieldType, attr.TransformMethod, rootContext.GetType());
                        field.SetValue(instance, value);
                    }
                }
                else if (!field.FieldType.IsPrimitive && !field.FieldType.IsEnum && field.FieldType != typeof(string) && !typeof(Object).IsAssignableFrom(field.FieldType))
                {
                    object subObj = field.GetValue(instance);

                    if (subObj != null)
                    {
                        ApplyFieldsRecursive(subObj, row, field.FieldType, rootContext, casing);
                    }
                }
            }
        }

        private static object ConvertValue(string columnName, string input, System.Type targetType, string transformMethod, System.Type contextType)
        {
            MethodInfo method = !string.IsNullOrEmpty(transformMethod) ?
                contextType.GetMethod(transformMethod, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic) :
                FindTransformMethodByColumn(contextType, columnName);

            if (method != null)
            {
                try
                {
                    return method.Invoke(null, new object[] { input });
                }
                catch (System.Exception e)
                {
                    Debug.LogError($"Failed to convert column '{columnName}' input '{input}' to {targetType.Name}: {e.Message}");
                    return targetType.IsValueType ? System.Activator.CreateInstance(targetType) : null;
                }
            }

            try
            {
                return System.Convert.ChangeType(input, targetType);
            }
            catch (System.Exception e)
            {
                Debug.LogError($"Failed to convert column '{columnName}' input '{input}' to {targetType.Name}: {e.Message}");
                return targetType.IsValueType ? System.Activator.CreateInstance(targetType) : null;
            }
        }

        private static MethodInfo FindTransformMethodByColumn(System.Type contextType, string columnName)
        {
            MethodInfo[] methods = contextType.GetMethods(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public);

            MethodInfo method = methods.FirstOrDefault(m => m.GetCustomAttribute<SheetTransformAttribute>()?.ColumnName == columnName);

            if (method != null)
            {
                return method;
            }

            foreach (FieldInfo field in contextType.GetFields(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public))
            {
                if (field.FieldType.IsPrimitive || field.FieldType.IsEnum || field.FieldType == typeof(string) || typeof(Object).IsAssignableFrom(field.FieldType))
                {
                    continue;
                }

                method = FindTransformMethodByColumn(field.FieldType, columnName);

                if (method != null)
                {
                    return method;
                }
            }

            return null;
        }

        private static IEnumerable<System.Type> GetAllScriptableObjectSheetTypes()
        {
            return System.AppDomain.CurrentDomain.GetAssemblies().SelectMany((Assembly a) => a.GetTypes()).Where((System.Type t) => typeof(ScriptableObject).IsAssignableFrom(t) && t.GetCustomAttribute<SheetAttribute>() != null);
        }

        private static ScriptableObject FindScriptableObjectByName(System.Type type, string name)
        {
            string[] guids = AssetDatabase.FindAssets($"t:{type.Name}");

            foreach (string guid in guids)
            {
                string path = AssetDatabase.GUIDToAssetPath(guid);
                ScriptableObject so = AssetDatabase.LoadAssetAtPath(path, type) as ScriptableObject;

                if (so.name == name)
                {
                    return so;
                }
            }

            return null;
        }
    }
}
