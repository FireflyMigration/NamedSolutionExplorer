using NamedSolutionExplorer.Models;

using System;
using System.Collections.Generic;
using System.Linq;

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