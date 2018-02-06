// /**********************************************************************************************
// Author:  Vasily Kabanov
// Created  2018-02-06
// Comment  
// **********************************************************************************************/

using System.Linq.Expressions;
using System.Reflection;

namespace ExcelEi.Read
{
    public static class ExpressionHelper
    {
        /// <summary>
        ///     Get property or field info.
        /// </summary>
        /// <param name="expression">
        ///     Simple reference, arbitrary number of conversions allowed.
        /// </param>
        /// <returns>
        ///     Null if not found
        /// </returns>
        public static MemberInfo GetMember(LambdaExpression expression)
        {
            var memberExpression = expression.Body as MemberExpression;
            if (memberExpression != null)
                return memberExpression.Member;

            var unaryExpression = expression.Body as UnaryExpression;
            Expression unaryOperandExpression = null;

            while (unaryExpression != null)
            {
                unaryOperandExpression = unaryExpression.Operand;
                unaryExpression = unaryOperandExpression as UnaryExpression;
            }

            return (unaryOperandExpression as MemberExpression)?.Member;
        }
    }
}