using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace NamedSolutionExplorer
{
    public class NSESettings
    {
        public class aNSE
        {
            public aNSE()
            {
            }

            public aNSE(string hierarchyId, string name) : this()
            {
                HierarchyId = hierarchyId;
                Name = name;
            }

            public string HierarchyId { get; set; }
            public string Name { get; set; }
        }

        public static async Task<NSESettings> FromString(string src)
        {
            return await Task.Run(() => JsonConvert.DeserializeObject<NSESettings>(src));
        }

        public List<aNSE> Settings { get; set; } = new List<aNSE>();

        public void Add(aNSE src)
        {
            Settings.Add(src);
        }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}