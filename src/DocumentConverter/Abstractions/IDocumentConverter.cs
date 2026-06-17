namespace DocumentConverter.Abstractions
{
	using System.IO;
	using DocumentConverter.Models;

	/// <summary>
	/// Defines a contract for converting a document stream to HTML format.
	/// </summary>
	public interface IDocumentConverter
	{
		/// <summary>
		/// Converts the given document stream to an HTML string.
		/// </summary>
		/// <param name="stream">The stream containing the document data.</param>
		/// <returns>A <see cref="Result{String}"/> containing the HTML string on success, or an error message on failure.</returns>
		Result<string> Convert(Stream stream);
	}
}
