using System.ComponentModel.DataAnnotations;

namespace CvBuilderWebV2.Dtos
{
    public sealed class CVFieldValuesDto
    {
        [Required(ErrorMessage = "Ime je obavezno")]
        public string FirstName { get; set; }

        [Required(ErrorMessage = "Prezime je obavezno")]
        public string LastName { get; set; }

        [Required(ErrorMessage = "Pozicija je obavezna")]
        public string Title { get; set; }

        [Required(ErrorMessage = "Ovaj deo ne sme biti prazan")]
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
