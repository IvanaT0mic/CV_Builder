namespace CVBuilder.Dtos
{
    public sealed class ExperienceDto
    {
        public string? ExpPosition { get; set; }
        public string? ExpStartDate { get; set; }
        public string? ExpEndDate { get; set; }
        public string? ExpCompanyName { get; set; }
        public string? ExpProjectName { get; set; }
        public string? ExpProjectDescription { get; set; }
        public string? ExpTechStack { get; set; }
        public List<string> ExpResponsibilities { get; set; } = [];
    }
}
