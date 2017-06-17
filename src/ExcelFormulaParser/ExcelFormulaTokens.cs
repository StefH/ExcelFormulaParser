using System.Collections.Generic;

namespace ExcelFormulaParser
{
    internal class ExcelFormulaTokens
    {
        private int _index = -1;
        private readonly List<ExcelFormulaToken> _tokens = new List<ExcelFormulaToken>();

        public int Count => _tokens.Count;

        public bool BOF => _index <= 0;

        public bool EOF => _index >= _tokens.Count - 1;

        public ExcelFormulaToken Current
        {
            get
            {
                if (_index == -1)
                {
                    return null;
                }

                return _tokens[_index];
            }
        }

        public ExcelFormulaToken Next
        {
            get
            {
                if (EOF)
                {
                    return null;
                }

                return _tokens[_index + 1];
            }
        }

        public ExcelFormulaToken Previous
        {
            get
            {
                if (_index < 1)
                {
                    return null;
                }

                return _tokens[_index - 1];
            }
        }

        public ExcelFormulaToken Add(ExcelFormulaToken token)
        {
            _tokens.Add(token);
            return token;
        }

        public bool MoveNext()
        {
            if (EOF)
            {
                return false;
            }

            _index++;
            return true;
        }

        public void Reset()
        {
            _index = -1;
        }
    }
}