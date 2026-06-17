namespace DocumentConverter.Models
{
	/// <summary>
	/// Represents the result of an operation, indicating success or failure.
	/// Used to avoid throwing exceptions for expected business errors.
	/// </summary>
	/// <typeparam name="T">The type of the success value.</typeparam>
	public class Result<T>
	{
		/// <summary>
		/// Gets a value indicating whether the operation was successful.
		/// </summary>
		public bool IsSuccess { get; }

		/// <summary>
		/// Gets the value returned by a successful operation.
		/// </summary>
		public T Value { get; }

		/// <summary>
		/// Gets the error message of a failed operation.
		/// </summary>
		public string ErrorMessage { get; }

		/// <summary>
		/// Initializes a new instance of the <see cref="Result{T}"/> class.
		/// </summary>
		/// <param name="isSuccess">True if successful; otherwise, false.</param>
		/// <param name="value">The success value.</param>
		/// <param name="errorMessage">The error message.</param>
		private Result(bool isSuccess, T value, string errorMessage)
		{
			IsSuccess = isSuccess;
			Value = value;
			ErrorMessage = errorMessage;
		}

		/// <summary>
		/// Creates a successful result containing a value.
		/// </summary>
		/// <param name="value">The value of the successful operation.</param>
		/// <returns>A successful <see cref="Result{T}"/> instance.</returns>
		public static Result<T> Success(T value)
		{
			return new Result<T>(true, value, null);
		}

		/// <summary>
		/// Creates a failed result containing an error message.
		/// </summary>
		/// <param name="errorMessage">The error message describing the failure.</param>
		/// <returns>A failed <see cref="Result{T}"/> instance.</returns>
		public static Result<T> Failure(string errorMessage)
		{
			return new Result<T>(false, default, errorMessage);
		}
	}
}
