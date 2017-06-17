using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace ExcelFormulaExpressionParser.Expressions
{
    internal class XArgExpression : Expression
    {
        private readonly List<Expression> _expressions;

        public IEnumerable<Expression> Expressions => _expressions;

        private XArgExpression(Expression expression)
        {
            _expressions = new[] { expression }.ToList();
        }

        public static XArgExpression Create(Expression expression)
        {
            return new XArgExpression(expression);
        }

        public XArgExpression Add(Expression expression)
        {
            _expressions.Add(expression);
            return this;
        }

        public override Type Type => typeof(IList<Expression>);

        public sealed override ExpressionType NodeType => ExpressionType.Constant;
    }
}