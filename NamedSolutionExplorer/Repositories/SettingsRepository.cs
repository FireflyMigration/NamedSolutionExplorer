using System;
using System.Collections.Generic;
using System.Linq;
using NamedSolutionExplorer.Models;

namespace NamedSolutionExplorer.Repositories
{
    public class SettingsRepository
    {
        #region Private Vars

        private readonly List<NamedSolutionExplorerWindowConfig> _windowConfigs =
            new List<NamedSolutionExplorerWindowConfig>();

        #endregion

        #region Properties

        public IEnumerable<NamedSolutionExplorerWindowConfig> WindowConfigs => _windowConfigs.AsEnumerable();

        #endregion

        #region Public Methods

        public void AddOrReplace(NamedSolutionExplorerWindowConfig config)
        {
            _windowConfigs.RemoveAll(x => x.Name.Equals(config.Name, StringComparison.InvariantCultureIgnoreCase));

            _windowConfigs.Add(config);
        }

        #endregion

        #region Nested Types

        public class SettingStorage
        {
            #region Properties

            public NamedSolutionExplorerWindowConfig[] Settings { get; set; }

            #endregion
        }

        #endregion
    }
}