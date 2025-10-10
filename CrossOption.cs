using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace FilterV1
{
    /// <summary>
    /// Represents a single custom cross option mapping.  Each option has an Id used for
    /// keyboard shortcuts and persistence, a human readable label displayed to the user,
    /// and up to three column values that will be applied when this option is selected.
    /// Empty strings are allowed for the column values in order to support blank
    /// hylsetyper.
    /// </summary>
    public class CrossOption
    {
        /// <summary>Numeric identifier for the option.  Values start at 1 and are used
        /// when handling keyboard shortcuts.  The Id must be unique within the list of
        /// options.  Option 0 is reserved for a blank entry and is not persisted.</summary>
        public int Id { get; set; }

        /// <summary>User visible label for the option.  This is displayed in drop downs
        /// and context menus within the custom cross section window.</summary>
        public string Label { get; set; } = string.Empty;

        /// <summary>Value assigned to column 2 of the target table when this option is
        /// applied.  May be an empty string to leave the cell blank.</summary>
        public string Col2 { get; set; } = string.Empty;

        /// <summary>Value assigned to column 3 of the target table when this option is
        /// applied.  May be an empty string to leave the cell blank.</summary>
        public string Col3 { get; set; } = string.Empty;

        /// <summary>Value assigned to column 4 of the target table when this option is
        /// applied.  May be an empty string to leave the cell blank.</summary>
        public string Col4 { get; set; } = string.Empty;
    }

    /// <summary>
    /// Provides loading and saving of cross options from a JSON file under the user's
    /// application data folder.  This static helper centralizes the file location and
    /// ensures defaults are created when no file exists.
    /// </summary>
    public static class CrossOptionRepository
    {
        private static string GetOptionsFilePath()
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FilterV1");
            Directory.CreateDirectory(folder);
            return Path.Combine(folder, "cross_options.json");
        }

        /// <summary>
        /// Loads the list of cross options from disk.  If no file exists, a default set
        /// of options is returned and persisted.  Any deserialization errors will also
        /// cause the defaults to be returned.
        /// </summary>
        public static List<CrossOption> Load()
        {
            try
            {
                string path = GetOptionsFilePath();
                if (File.Exists(path))
                {
                    string json = File.ReadAllText(path);
                    var loaded = JsonSerializer.Deserialize<List<CrossOption>>(json);
                    if (loaded != null && loaded.Count > 0)
                    {
                        // Ensure Ids are unique and greater than 0.  Remove invalid entries.
                        var distinct = loaded
                            .Where(opt => opt != null && opt.Id > 0)
                            .GroupBy(opt => opt.Id)
                            .Select(g => g.First())
                            .OrderBy(opt => opt.Id)
                            .ToList();
                        if (distinct.Count > 0)
                        {
                            return distinct;
                        }
                    }
                }
            }
            catch
            {
                // ignore errors and fall back to defaults
            }
            // If loading fails, return defaults and save them to disk
            var defaults = GetDefaultOptions();
            Save(defaults);
            return defaults;
        }

        /// <summary>
        /// Saves the provided list of options to disk.  Invalid or duplicate Ids are
        /// filtered out automatically.  Returns true if the operation succeeds.
        /// </summary>
        public static bool Save(List<CrossOption> options)
        {
            try
            {
                var valid = options
                    .Where(opt => opt != null && opt.Id > 0)
                    .GroupBy(opt => opt.Id)
                    .Select(g => g.First())
                    .OrderBy(opt => opt.Id)
                    .ToList();
                string path = GetOptionsFilePath();
                string json = JsonSerializer.Serialize(valid, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Returns a default list of cross options similar to the previous hard coded
        /// values.  These defaults cover common wire sizes and one RDX4K6.0 case.
        /// </summary>
        public static List<CrossOption> GetDefaultOptions()
        {
            return new List<CrossOption>
            {
                new CrossOption { Id = 1, Label = "1.0 mm²",  Col2 = "UNIBK1.0",  Col3 = "HYLSE 1.0",  Col4 = "HYLSE 1.0" },
                new CrossOption { Id = 2, Label = "1.5 mm²",  Col2 = "UNIBK1.5",  Col3 = "HYLSE 1.5",  Col4 = "HYLSE 1.5" },
                new CrossOption { Id = 3, Label = "2.5 mm²",  Col2 = "UNIBK2.5",  Col3 = "HYLSE 2.5",  Col4 = "HYLSE 2.5" },
                new CrossOption { Id = 4, Label = "4.0 mm²",  Col2 = "UNIBK4.0",  Col3 = "HYLSE 4.0",  Col4 = "HYLSE 4.0" },
                new CrossOption { Id = 5, Label = "6.0 mm² (RDX4K6.0)",  Col2 = "RDX4BK6.0", Col3 = string.Empty,      Col4 = string.Empty }
            };
        }
    }
}