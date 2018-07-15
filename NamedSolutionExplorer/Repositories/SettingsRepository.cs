using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using NamedSolutionExplorer.Models;
using Newtonsoft.Json;

namespace NamedSolutionExplorer.Repositories
{
    public class SettingsRepository
    {
        #region Private Vars

        private List<NamedSolutionExplorerWindowConfig> _windowConfigs =
            new List<NamedSolutionExplorerWindowConfig>();

        #endregion Private Vars

        #region Properties

        public IEnumerable<NamedSolutionExplorerWindowConfig> WindowConfigs => _windowConfigs.AsEnumerable();

        #endregion Properties

        #region Public Methods

        public void AddOrReplace(NamedSolutionExplorerWindowConfig config)
        {
            _windowConfigs.RemoveAll(x => x.Name.Equals(config.Name, StringComparison.InvariantCultureIgnoreCase));

            _windowConfigs.Add(config);
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

        public async Task Save(string filePath)
        {
            using (var f = File.CreateText(filePath))
            {
                var storedForm = new SettingStorage { Settings = _windowConfigs.ToArray() };
                await f.WriteAsync(JsonConvert.SerializeObject(storedForm, Formatting.Indented));
            }
        }

        #endregion Public Methods

        #region Nested Types

        public class SettingStorage
        {
            #region Properties

            public NamedSolutionExplorerWindowConfig[] Settings { get; set; }

            #endregion Properties
        }

        #endregion Nested Types
    }
}