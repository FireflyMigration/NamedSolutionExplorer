using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace NamedSolutionExplorer
{
    public class SettingsRepository
    {
        public class SettingStorage
        {
            public NamedSolutionExplorerWindowConfig[] Settings { get; set; }
        }

        private List<NamedSolutionExplorerWindowConfig> _windowConfigs =
            new List<NamedSolutionExplorerWindowConfig>();

        public IEnumerable<NamedSolutionExplorerWindowConfig> WindowConfigs => _windowConfigs.AsEnumerable();

        public async Task Save(string filePath)
        {
            using (var f = File.CreateText(filePath))
            {
                var storedForm = new SettingStorage() { Settings = _windowConfigs.ToArray() };
                await f.WriteAsync(JsonConvert.SerializeObject(storedForm, Formatting.Indented));
            }
        }

        public async Task<bool> Load(string filePath)
        {
            using (var f = File.OpenText(filePath))
            {
                var contents = await f.ReadToEndAsync();
                var storedForm = JsonConvert.DeserializeObject<SettingStorage>(contents);

                _windowConfigs = new List<NamedSolutionExplorerWindowConfig>(storedForm.Settings);

                return true;
            }
        }

        public void AddOrReplace(NamedSolutionExplorerWindowConfig config)
        {
            _windowConfigs.RemoveAll(x => x.Name.Equals(config.Name, StringComparison.InvariantCultureIgnoreCase));

            _windowConfigs.Add(config);
        }
    }
}