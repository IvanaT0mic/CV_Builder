using Colorbook.Shared.V2.Models;
using Microsoft.AspNetCore.Http.Features;
using System.Net;

namespace CVBuilder.Middlewares
{
    /// <summary>
    /// Exception middleware
    /// </summary>
    public sealed class ExceptionMiddleware(
        RequestDelegate _next,
        ILogger<ExceptionMiddleware> _logger
    )
    {
        /// <summary>
        /// Invoke method
        /// </summary>
        /// <param name="context">Http context</param>
        /// <returns></returns>
        public async Task InvokeAsync(HttpContext context)
        {
            try
            {
                await _next(context);
            }
            catch (Exception exception)
            {
                context.Response.ContentType = "text/javascript";
                var contextResponse = context.Features.Get<IHttpResponseFeature>();

                if (contextResponse == null)
                {
                    await context.Response.WriteAsync("Cannot initialize response context");
                    return;
                }

                if (exception is BadRequestError)
                {
                    _logger.LogWarning(
                        "{path} 400 BadRequest - {msg}",
                        context.Request.Path,
                        exception.Message
                    );

                    contextResponse.StatusCode = (int)HttpStatusCode.BadRequest;
                    await context.Response.WriteAsync(exception.Message);
                }
                else if (exception is UnauthorizedError)
                {
                    _logger.LogWarning(
                        "{path} 401 Unauthorized - {msg}",
                        context.Request.Path,
                        exception.Message
                    );

                    contextResponse.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await context.Response.WriteAsync(exception.Message);
                }
                else if (exception is UnauthorizedAccessException)
                {
                    _logger.LogWarning(
                        "{path} 401 Unauthorized - {msg}",
                        context.Request.Path,
                        exception.Message
                    );
                    contextResponse.StatusCode = (int)HttpStatusCode.Unauthorized;
                    await context.Response.WriteAsync(exception.Message);
                }
                else
                {
                    _logger.LogError(
                        "{path} 500 InternalServerError - {msg}",
                        context.Request.Path,
                        string.Concat(
                            exception.Message,
                            " ",
                            exception.InnerException?.Message ?? string.Empty
                        )
                    );
                    contextResponse.StatusCode = (int)HttpStatusCode.InternalServerError;
                    await context.Response.WriteAsync(exception.Message);
                }
            }
        }
    }
}
