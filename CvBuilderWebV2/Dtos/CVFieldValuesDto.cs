namespace CvBuilderWebV2.Dtos
{
    public sealed class CVFieldValuesDto
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Title { get; set; }
        public string AboutMe { get; set; }
        public string Skills { get; set; }
        public string ToolStack { get; set; }
        public string Languages { get; set; }
        public List<ProfileGradesDto> ProfileGrades { get; set; }
        public List<ExperienceDto> Experiences { get; set; }
        public List<TrainingDto> Trainings { get; set; }
        public List<SocialDto> Socials { get; set; }
        public List<EducationDto> Educations { get; set; }
    }
}
