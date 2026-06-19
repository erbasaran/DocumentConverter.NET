namespace DocumentConverter.Abstractions
{
	using DocumentConverter.Models;

	/// <summary>
	/// Defines a contract for converting an HTML string to a binary document format.
	/// </summary>
	public interface IHtmlToDocumentConverter
	{
		/// <summary>
		/// Converts the given HTML string to a byte array representation of the target document.
		/// </summary>
		/// <param name="html">The HTML string content.</param>
		/// <returns>A <see cref="Result{Byte[]}"/> containing the output bytes on success, or an error message on failure.</returns>
		Result<byte[]> Convert(string html);
	}
}
