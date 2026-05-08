namespace CVBuilder.Dtos
{
    public sealed class CVFieldValuesDto
    {
        public required string FirstName { get; set; }
        public required string LastName { get; set; }
        public required string Title { get; set; }
        public required string AboutMe { get; set; }
        public List<ProfileGradesDto> ProfileGrades { get; set; }
        public string Skills { get; set; }
        public List<ExperienceDto> Experiences { get; set; }
        public List<TrainingDto> Trainings { get; set; }
        public List<SocialDto> Socials { get; set; }
        public List<EducationDto> Educations { get; set; }
        public string ToolStack { get; set; }
        public string Languages { get; set; }
    }
}
