namespace NamedSolutionExplorer
{
    public class NamedSolutionExplorerWindowConfig
    {
        public NamedSolutionExplorerWindowConfig(string hierarchyId, string name)
        {
            HierarchyId = hierarchyId;
            Name = name;
        }

        public string HierarchyId { get; set; }
        public string Name { get; set; }
    }
}