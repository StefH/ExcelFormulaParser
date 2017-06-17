using System;
using System.Collections.Generic;

namespace ExcelFormulaParser
{
    internal class ExcelFormulaStack
    {
        private readonly Stack<ExcelFormulaToken> _stack = new Stack<ExcelFormulaToken>();

        public ExcelFormulaToken Current => _stack.Count > 0 ? _stack.Peek() : null;

        public void Push(ExcelFormulaToken token)
        {
            _stack.Push(token);
        }

        public ExcelFormulaToken Pop()
        {
            if (_stack.Count == 0)
            {
                return null;
            }

            return new ExcelFormulaToken(String.Empty, _stack.Pop().Type, ExcelFormulaTokenSubtype.Stop);
        }
    }
}