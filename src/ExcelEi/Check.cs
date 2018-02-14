using System;

namespace ExcelEi
{
    public static class Check
    {
        /// <summary>
        ///		Throws ArgumentNullException if necessary.
        /// </summary>
        public static void DoRequireArgumentNotNull(object arg, string argName)
        {
            if (arg == null)
                throw new ArgumentNullException(argName);
        }

        /// <summary>
        ///		Throws ArgumentException if necessary.
        /// </summary>
        public static void DoRequireArgumentNotBlank(string arg, string argName)
        {
            if (string.IsNullOrWhiteSpace(arg))
                throw new ArgumentException($"{argName} must not be blank");
        }

        public static void DoRequire(bool assertion, string message)
        {
            if (!assertion) throw new PreconditionException(message);
        }

        public static void DoEnsure(bool assertion, string message)
        {
            if (!assertion) throw new PostconditionException(message);
        }

        public static void DoEnsureLambda(bool assertion, Func<string> message)
        {
            if (!assertion) throw new PostconditionException(message.Invoke());
        }

        public static void DoAssertLambda(bool assertion, Func<string> message)
        {
            if (!assertion) throw new AssertionException(message.Invoke());
        }

        public static void DoRequire(bool assertion, string message, Exception inner)
        {
            if (!assertion) throw new PreconditionException(message, inner);
        }

        public static void DoAssertLambda(bool assertion, Func<string> message, Func<Exception> inner)
        {
            if (!assertion) throw new PreconditionException(message.Invoke(), inner.Invoke());
        }

        public static void DoAssertLambda(bool assertion, Func<Exception> getExceptionToThrow)
        {
            if (!assertion) throw getExceptionToThrow.Invoke();
        }

        public static void DoCheckArgument(bool assertion, Func<string> getErrorMessage)
        {
            if (!assertion)
                throw new ArgumentException(getErrorMessage());
        }

        public static void DoCheckArgument(bool assertion, string message = "", string argName = "")
        {
            if (!assertion)
                throw new ArgumentException(message, argName);
        }


        /// <summary>
        ///		<see cref="InvalidOperationException"/> thrown if assertion fails
        /// </summary>
        /// <param name="assertion">
        ///		The assertion condition
        /// </param>
        /// <param name="exceptionMessage">
        ///		Message to return from the thrown exception if assertion fails. Pass <see langword="null"/>
        ///		to use default constructor for <see cref="InvalidOperationException"/>
        /// </param>
        public static void DoCheckOperationValid(bool assertion, string exceptionMessage)
        {
            if (assertion) return;

            if (string.IsNullOrEmpty(exceptionMessage))
                throw new InvalidOperationException();

            throw new InvalidOperationException(exceptionMessage);
        }

        /// <summary>
        ///		<see cref="InvalidOperationException"/> thrown if assertion fails
        /// </summary>
        /// <param name="assertion">
        ///		The assertion condition
        /// </param>
        /// <param name="getExceptionMessage">
        ///		Delegate retrieving message to return from the thrown exception if assertion fails. Pass <see langword="null"/>
        ///		to use default constructor for <see cref="InvalidOperationException"/>
        /// </param>
        public static void DoCheckOperationValid(bool assertion, Func<string> getExceptionMessage)
        {
            if (assertion) return;

            if (getExceptionMessage == null)
            {
                throw new InvalidOperationException();
            }

            throw new InvalidOperationException(getExceptionMessage());
        }

        /// <summary>
        ///		<see cref="InvalidOperationException"/> thrown if assertion fails
        /// </summary>
        /// <param name="assertion">
        ///		The assertion condition
        /// </param>
        /// <remarks>
        ///		The exception if thrown is instantiated using the default constructor
        /// </remarks>
        public static void DoCheckOperationValid(bool assertion)
        {
            DoCheckOperationValid(assertion, (string) null);
        }

        public static void DoCheckInvariant(bool assertion, string exceptionMessage)
        {
            if (assertion) return;

            throw new InvariantException(exceptionMessage);
        }

        public static void DoCheckInvariant(bool assertion, Func<string> exceptionMessageGetter)
        {
            if (assertion) return;

            throw new InvariantException(exceptionMessageGetter?.Invoke());
        }

        public static void DoRequire(bool assertion)
        {
            if (!assertion) throw new PreconditionException("Precondition failed.");
        }
    }

    public class DesignByContractException : Exception
    {
        protected DesignByContractException()
        {
        }

        protected DesignByContractException(string message)
            : base(message)
        {
        }

        protected DesignByContractException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    [Serializable]
    public class PreconditionException : DesignByContractException
    {
        public PreconditionException()
        {
        }

        public PreconditionException(string message)
            : base(message)
        {
        }

        public PreconditionException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    [Serializable]
    public class PostconditionException : DesignByContractException
    {
        public PostconditionException()
        {
        }

        public PostconditionException(string message) : base(message)
        {
        }

        public PostconditionException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    [Serializable]
    public class InvariantException : DesignByContractException
    {
        public InvariantException()
        {
        }

        public InvariantException(string message)
            : base(message)
        {
        }

        public InvariantException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    [Serializable]
    public class AssertionException : DesignByContractException
    {
        public AssertionException()
        {
        }

        public AssertionException(string message)
            : base(message)
        {
        }

        public AssertionException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}