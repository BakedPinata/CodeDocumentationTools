namespace CodeDocumenter.Shared;

/// <summary>
/// What a tool documents can't be read from what it was given: it isn't the kind of file or folder the tool reads, it
/// doesn't hold what the tool reads, it's encrypted or it's damaged. The message says which, in words the user can act
/// on. The run stops with exit code 3.
/// </summary>
public sealed class ExtractionException : Exception
{
  /// <summary>Creates the exception with a message for the user.</summary>
  public ExtractionException(string message)
    : base(message)
  {
  }

  /// <summary>Creates the exception with a message for the user and the error behind it.</summary>
  public ExtractionException(string message, Exception? innerException)
    : base(message, innerException)
  {
  }
}
