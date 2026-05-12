namespace Colorbook.Shared.V2.Models
{
    public class UnauthorizedError : Exception
    {
        public UnauthorizedError(string message) : base(message)
        {
        }
    }
}
