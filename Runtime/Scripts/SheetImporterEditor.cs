using HHG.GoogleSheets.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using UnityEditor;
using UnityEngine;

namespace HHG.GoogleSheets.Editor
{
    public class SheetImporterEditor : EditorWindow
    {
        [MenuItem("Tools/Import All Sheets")]
        public static void ImportAllSheets()
        {
            foreach (var type in GetAllSheetTypes())
            {
                var attr = type.GetCustomAttribute<SheetAttribute>();
                if (attr == null) continue;

                Debug.Log($"Importing sheet for {type.Name}: {attr.Gid}");

                string csv = DownloadSheetCSV(attr.SpreadsheetId, attr.Gid);
                if (string.IsNullOrEmpty(csv))
                {
                    Debug.LogError($"Failed to download sheet for {type.Name}");
                    continue;
                }

                ImportCSVToScriptableObjects(csv, type);
            }

            AssetDatabase.SaveAssets();
            Debug.Log("✅ All sheets imported!");
        }

        static string DownloadSheetCSV(string spreadsheetId, string gid)
        {
            try
            {
                string url = $"https://docs.google.com/spreadsheets/d/{spreadsheetId}/export?format=csv";
                if (!string.IsNullOrEmpty(gid))
                {
                    url += $"&gid={Uri.EscapeDataString(gid)}";
                }

                using WebClient client = new WebClient();
                return client.DownloadString(url);
            }
            catch (Exception ex)
            {
                Debug.LogError($"Error downloading Google Sheet: {ex.Message}");
                return null;
            }
        }

        static void ImportCSVToScriptableObjects(string csv, Type soType)
        {
            var lines = csv.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length < 2) return;

            var headers = lines[0].Split(',').Select(h => h.Trim()).ToArray();

            for (int i = 1; i < lines.Length; i++)
            {
                var values = lines[i].Split(',').Select(v => v.Trim()).ToArray();
                var row = headers.Zip(values, (k, v) => new { k, v }).ToDictionary(x => x.k, x => x.v);

                if (!row.TryGetValue("Name", out string name) || string.IsNullOrEmpty(name)) continue;

                var so = FindSOByName(soType, name);
                if (so == null)
                {
                    Debug.LogWarning($"No ScriptableObject found for {soType.Name} named '{name}'");
                    continue;
                }

                ApplyFieldsRecursive(so, row, soType, so);
                EditorUtility.SetDirty(so);
            }
        }

        static void ApplyFieldsRecursive(object instance, Dictionary<string, string> row, Type type, object rootContext)
        {
            foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                var attr = field.GetCustomAttribute<SheetFieldAttribute>();
                if (attr != null)
                {
                    if (row.TryGetValue(attr.ColumnName, out string raw))
                    {
                        object value = ConvertValue(raw, field.FieldType, attr.TransformMethod, rootContext.GetType());
                        field.SetValue(instance, value);
                    }
                }
                else if (!field.FieldType.IsPrimitive && !field.FieldType.IsEnum && field.FieldType != typeof(string))
                {
                    object subObj = field.GetValue(instance);
                    if (subObj != null)
                    {
                        ApplyFieldsRecursive(subObj, row, field.FieldType, rootContext);
                    }
                }
            }
        }

        static object ConvertValue(string raw, Type targetType, string transformMethod, Type contextType)
        {
            if (!string.IsNullOrEmpty(transformMethod))
            {
                var method = contextType.GetMethod(transformMethod, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                if (method != null)
                {
                    return method.Invoke(null, new object[] { raw });
                }
                else
                {
                    Debug.LogWarning($"Transformer method '{transformMethod}' not found on {contextType.Name}.");
                }
            }

            try
            {
                return Convert.ChangeType(raw, targetType);
            }
            catch
            {
                Debug.LogWarning($"Failed to convert '{raw}' to {targetType.Name}");
                return targetType.IsValueType ? Activator.CreateInstance(targetType) : null;
            }
        }


        static IEnumerable<Type> GetAllSheetTypes()
        {
            return AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(a => a.GetTypes())
                .Where(t => typeof(ScriptableObject).IsAssignableFrom(t)
                    && t.GetCustomAttribute<SheetAttribute>() != null);
        }

        static ScriptableObject FindSOByName(Type type, string name)
        {
            string[] guids = AssetDatabase.FindAssets($"t:{type.Name}");
            foreach (string guid in guids)
            {
                string path = AssetDatabase.GUIDToAssetPath(guid);
                var so = AssetDatabase.LoadAssetAtPath(path, type) as ScriptableObject;

                if (so.name == name)
                    return so;
            }

            return null;
        }
    } 
}
