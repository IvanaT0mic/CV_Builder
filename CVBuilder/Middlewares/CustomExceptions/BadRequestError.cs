namespace Colorbook.Shared.V2.Models
{
    public class BadRequestError : Exception
    {
        public BadRequestError(string message, params object?[] args) : base(string.Format(message, args)) { }
        public BadRequestError(string message) : base(message) { }
    }
}
