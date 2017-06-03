using JetBrains.Annotations;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser
{
    public class ExcelFormula : IList<ExcelFormulaToken>
    {
        const char QuoteDouble = '"';
        const char QuoteSingle = '\'';
        const char BracketClose = ']';
        const char BracketOpen = '[';
        const char BraceOpen = '{';
        const char BraceClose = '}';
        const char ParenOpen = '(';
        const char ParenClose = ')';
        const char Semicolon = ';';
        const char Whitespace = ' ';
        const char Comma = ',';
        const char ErrorStart = '#';

        const string OperatorsSn = "+-";
        const string OperatorsInfix = "+-*/^&=><";
        const string OperatorsPostfix = "%";

        internal readonly string[] ExcelErrors = { "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A" };
        internal readonly string[] ComparatorsMulti = { ">=", "<=", "<>" };

        private readonly string _formula;
        private List<ExcelFormulaToken> _tokens = new List<ExcelFormulaToken>();

        /// <summary>
        /// Gets the number of ExcelFormulaToken which are parsed for this ExcelFormula.
        /// </summary>
        public int Count => _tokens.Count;

        /// <summary>
        /// Gets a value indicating whether this ExcelFormula is read-only.
        /// </summary>
        public bool IsReadOnly => true;

        /// <summary>
        /// The ExcelFormula.
        /// </summary>
        public string Formula => _formula;

        /// <summary>
        /// The optional context for this formula, can be anything, e.g. the Sheet or Workbook.
        /// </summary>
        public IExcelFormulaContext Context { get; private set; }

        /// <summary>
        /// Gets or sets the ExcelFormulaToken at the specified index.
        /// </summary>
        /// <param name="index">The zero-based index of the element to get or set.</param>
        /// <returns>The ExcelFormulaToken at the specified index.</returns>
        public ExcelFormulaToken this[int index]
        {
            get => _tokens[index];
            set => throw new NotSupportedException();
        }

        /// <summary>
        /// Constructs a ExcelFormula.
        /// </summary>
        /// <param name="formula">The Excel formula.</param>
        /// <param name="context">The optional context (can be anything, e.g. the Sheet or Workbook)</param>
        public ExcelFormula([NotNull] string formula, [CanBeNull] IExcelFormulaContext context = null)
        {
            if (string.IsNullOrEmpty(formula))
            {
                throw new ArgumentException(nameof(formula));
            }

            _formula = formula.Trim();
            Context = context;

            ParseToTokens();
        }

        public int IndexOf(ExcelFormulaToken item)
        {
            return _tokens.IndexOf(item);
        }

        public void Insert(int index, ExcelFormulaToken item)
        {
            throw new NotSupportedException();
        }

        public void RemoveAt(int index)
        {
            throw new NotSupportedException();
        }

        public void Add(ExcelFormulaToken item)
        {
            throw new NotSupportedException();
        }

        public void Clear()
        {
            throw new NotSupportedException();
        }

        public bool Contains(ExcelFormulaToken item)
        {
            return _tokens.Contains(item);
        }

        public bool Remove(ExcelFormulaToken item)
        {
            throw new NotSupportedException();
        }

        public void CopyTo(ExcelFormulaToken[] array, int arrayIndex)
        {
            _tokens.CopyTo(array, arrayIndex);
        }

        public IEnumerator<ExcelFormulaToken> GetEnumerator()
        {
            return _tokens.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private void ParseToTokens()
        {
            // No attempt is made to verify formulas; assumes formulas are derived from Excel, where 
            // they can only exist if valid; stack overflows/underflows sunk as nulls without exceptions.
            if (_formula.Length < 2 || _formula[0] != '=')
            {
                return;
            }

            var tokens1 = new ExcelFormulaTokens();
            var stack = new ExcelFormulaStack();

            bool inString = false;
            bool inPath = false;
            bool inRange = false;
            bool inError = false;

            int index = 1;
            string value = string.Empty;

            while (index < _formula.Length)
            {
                // state-dependent character evaluation (order is important)
                // double-quoted strings
                // embeds are doubled
                // end marks token
                if (inString)
                {
                    if (_formula[index] == QuoteDouble)
                    {
                        if (index + 2 <= _formula.Length && _formula[index + 1] == QuoteDouble)
                        {
                            value += QuoteDouble;
                            index++;
                        }
                        else
                        {
                            inString = false;
                            tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand, ExcelFormulaTokenSubtype.Text));
                            value = string.Empty;
                        }
                    }
                    else
                    {
                        value += _formula[index];
                    }
                    index++;
                    continue;
                }

                // single-quoted strings (links)
                // embeds are double
                // end does not mark a token
                if (inPath)
                {
                    if (_formula[index] == QuoteSingle)
                    {
                        if (index + 2 <= _formula.Length && _formula[index + 1] == QuoteSingle)
                        {
                            value += QuoteSingle;
                            index++;
                        }
                        else
                        {
                            inPath = false;
                        }
                    }
                    else
                    {
                        value += _formula[index];
                    }
                    index++;
                    continue;
                }

                // bracked strings (R1C1 range index or linked workbook name)
                // no embeds (changed to "()" by Excel)
                // end does not mark a token
                if (inRange)
                {
                    if (_formula[index] == BracketClose)
                    {
                        inRange = false;
                    }
                    value += _formula[index];
                    index++;
                    continue;
                }

                // error values
                // end marks a token, determined from absolute list of values
                if (inError)
                {
                    value += _formula[index];
                    index++;
                    if (Array.IndexOf(ExcelErrors, value) != -1)
                    {
                        inError = false;
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand, ExcelFormulaTokenSubtype.Error));
                        value = string.Empty;
                    }
                    continue;
                }

                // scientific notation check
                if (OperatorsSn.IndexOf(_formula[index]) != -1)
                {
                    if (value.Length > 1)
                    {
                        if (Regex.IsMatch(value, @"^[1-9]{1}(\.[0-9]+)?E{1}$"))
                        {
                            value += _formula[index];
                            index++;
                            continue;
                        }
                    }
                }

                // independent character evaluation (order not important)
                // establish state-dependent character evaluations
                if (_formula[index] == QuoteDouble)
                {
                    if (value.Length > 0)
                    {
                        // unexpected
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Unknown));
                        value = string.Empty;
                    }
                    inString = true;
                    index++;
                    continue;
                }

                if (_formula[index] == QuoteSingle)
                {
                    if (value.Length > 0)
                    {
                        // unexpected
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Unknown));
                        value = string.Empty;
                    }
                    inPath = true;
                    index++;
                    continue;
                }

                if (_formula[index] == BracketOpen)
                {
                    inRange = true;
                    value += BracketOpen;
                    index++;
                    continue;
                }

                if (_formula[index] == ErrorStart)
                {
                    if (value.Length > 0)
                    {
                        // unexpected
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Unknown));
                        value = string.Empty;
                    }
                    inError = true;
                    value += ErrorStart;
                    index++;
                    continue;
                }

                // mark start and end of arrays and array rows
                if (_formula[index] == BraceOpen)
                {
                    if (value.Length > 0)
                    {
                        // unexpected
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Unknown));
                        value = string.Empty;
                    }
                    stack.Push(tokens1.Add(new ExcelFormulaToken("ARRAY", ExcelFormulaTokenType.Function, ExcelFormulaTokenSubtype.Start)));
                    stack.Push(tokens1.Add(new ExcelFormulaToken("ARRAYROW", ExcelFormulaTokenType.Function, ExcelFormulaTokenSubtype.Start)));
                    index++;
                    continue;
                }

                if (_formula[index] == Semicolon)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }
                    tokens1.Add(stack.Pop());
                    tokens1.Add(new ExcelFormulaToken(",", ExcelFormulaTokenType.Argument));
                    stack.Push(tokens1.Add(new ExcelFormulaToken("ARRAYROW", ExcelFormulaTokenType.Function, ExcelFormulaTokenSubtype.Start)));
                    index++;
                    continue;
                }

                if (_formula[index] == BraceClose)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }
                    tokens1.Add(stack.Pop());
                    tokens1.Add(stack.Pop());
                    index++;
                    continue;
                }

                // trim white-space
                if (_formula[index] == Whitespace)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }
                    tokens1.Add(new ExcelFormulaToken("", ExcelFormulaTokenType.Whitespace));
                    index++;
                    while (_formula[index] == Whitespace && index < _formula.Length)
                    {
                        index++;
                    }
                    continue;
                }

                // multi-character comparators
                if (index + 2 <= _formula.Length)
                {
                    if (Array.IndexOf(ComparatorsMulti, _formula.Substring(index, 2)) != -1)
                    {
                        if (value.Length > 0)
                        {
                            tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                            value = string.Empty;
                        }
                        tokens1.Add(new ExcelFormulaToken(_formula.Substring(index, 2), ExcelFormulaTokenType.OperatorInfix, ExcelFormulaTokenSubtype.Logical));
                        index += 2;
                        continue;
                    }
                }

                // standard infix operators
                if (OperatorsInfix.IndexOf(_formula[index]) != -1)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }
                    tokens1.Add(new ExcelFormulaToken(_formula[index].ToString(), ExcelFormulaTokenType.OperatorInfix));
                    index++;
                    continue;
                }

                // standard postfix operators (only one)
                if (OperatorsPostfix.IndexOf(_formula[index]) != -1)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }
                    tokens1.Add(new ExcelFormulaToken(_formula[index].ToString(), ExcelFormulaTokenType.OperatorPostfix));
                    index++;
                    continue;
                }

                // start subexpression or function
                if (_formula[index] == ParenOpen)
                {
                    if (value.Length > 0)
                    {
                        stack.Push(tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Function, ExcelFormulaTokenSubtype.Start)));
                        value = string.Empty;
                    }
                    else
                    {
                        stack.Push(tokens1.Add(new ExcelFormulaToken("", ExcelFormulaTokenType.Subexpression, ExcelFormulaTokenSubtype.Start)));
                    }
                    index++;
                    continue;
                }

                // function, subexpression, or array parameters, or operand unions
                if (_formula[index] == Comma)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }

                    if (stack.Current.Type != ExcelFormulaTokenType.Function)
                    {
                        tokens1.Add(new ExcelFormulaToken(",", ExcelFormulaTokenType.OperatorInfix, ExcelFormulaTokenSubtype.Union));
                    }
                    else
                    {
                        tokens1.Add(new ExcelFormulaToken(",", ExcelFormulaTokenType.Argument));
                    }
                    index++;
                    continue;
                }

                // stop subexpression
                if (_formula[index] == ParenClose)
                {
                    if (value.Length > 0)
                    {
                        tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
                        value = string.Empty;
                    }
                    tokens1.Add(stack.Pop());
                    index++;
                    continue;
                }

                // token accumulation
                value += _formula[index];
                index++;
            }

            // dump remaining accumulation
            if (value.Length > 0)
            {
                tokens1.Add(new ExcelFormulaToken(value, ExcelFormulaTokenType.Operand));
            }

            // move tokenList to new set, excluding unnecessary white-space tokens and converting necessary ones to intersections
            var tokens2 = new ExcelFormulaTokens();
            while (tokens1.MoveNext())
            {
                ExcelFormulaToken token = tokens1.Current;

                if (token == null)
                {
                    continue;
                }

                if (token.Type != ExcelFormulaTokenType.Whitespace)
                {
                    tokens2.Add(token);
                    continue;
                }

                if (tokens1.BOF || tokens1.EOF)
                {
                    continue;
                }

                ExcelFormulaToken previous = tokens1.Previous;

                if (previous == null)
                {
                    continue;
                }

                if (!(
                        previous.Type == ExcelFormulaTokenType.Function &&
                        previous.Subtype == ExcelFormulaTokenSubtype.Stop ||
                        previous.Type == ExcelFormulaTokenType.Subexpression &&
                        previous.Subtype == ExcelFormulaTokenSubtype.Stop ||
                        previous.Type == ExcelFormulaTokenType.Operand
                    ))
                {
                    continue;
                }

                ExcelFormulaToken next = tokens1.Next;

                if (next == null)
                {
                    continue;
                }

                if (!(
                        next.Type == ExcelFormulaTokenType.Function && next.Subtype == ExcelFormulaTokenSubtype.Start ||
                        next.Type == ExcelFormulaTokenType.Subexpression &&
                        next.Subtype == ExcelFormulaTokenSubtype.Start ||
                        next.Type == ExcelFormulaTokenType.Operand
                    )
                )
                {
                    continue;
                }

                tokens2.Add(new ExcelFormulaToken("", ExcelFormulaTokenType.OperatorInfix, ExcelFormulaTokenSubtype.Intersection));
            }

            // move tokens to final list, switching infix "-" operators to prefix when appropriate, switching infix "+" operators 
            // to noop when appropriate, identifying operand and infix-operator subtypes, and pulling "@" from function names
            _tokens = new List<ExcelFormulaToken>(tokens2.Count);
            while (tokens2.MoveNext())
            {
                ExcelFormulaToken token = tokens2.Current;

                if (token == null)
                {
                    continue;
                }

                ExcelFormulaToken previous = tokens2.Previous;
                ExcelFormulaToken next = tokens2.Next;

                if (token.Type == ExcelFormulaTokenType.OperatorInfix && token.Value == "-")
                {
                    if (tokens2.BOF)
                    {
                        token.Type = ExcelFormulaTokenType.OperatorPrefix;
                    }
                    else if (
                        previous.Type == ExcelFormulaTokenType.Function &&
                        previous.Subtype == ExcelFormulaTokenSubtype.Stop ||
                        previous.Type == ExcelFormulaTokenType.Subexpression &&
                        previous.Subtype == ExcelFormulaTokenSubtype.Stop ||
                        previous.Type == ExcelFormulaTokenType.OperatorPostfix ||
                        previous.Type == ExcelFormulaTokenType.Operand
                    )
                    {
                        token.Subtype = ExcelFormulaTokenSubtype.Math;
                    }
                    else
                    {
                        token.Type = ExcelFormulaTokenType.OperatorPrefix;
                    }

                    _tokens.Add(token);
                    continue;
                }

                if (token.Type == ExcelFormulaTokenType.OperatorInfix && token.Value == "+")
                {
                    if (tokens2.BOF)
                    {
                        continue;
                    }

                    if (
                        previous.Type == ExcelFormulaTokenType.Function &&
                        previous.Subtype == ExcelFormulaTokenSubtype.Stop ||
                        previous.Type == ExcelFormulaTokenType.Subexpression &&
                        previous.Subtype == ExcelFormulaTokenSubtype.Stop ||
                        previous.Type == ExcelFormulaTokenType.OperatorPostfix ||
                        previous.Type == ExcelFormulaTokenType.Operand
                    )
                    {
                        token.Subtype = ExcelFormulaTokenSubtype.Math;
                    }
                    else
                    {
                        continue;
                    }

                    _tokens.Add(token);
                    continue;
                }

                if (token.Type == ExcelFormulaTokenType.OperatorInfix && token.Subtype == ExcelFormulaTokenSubtype.Nothing)
                {
                    if ("<>=".IndexOf(token.Value.Substring(0, 1), StringComparison.Ordinal) != -1)
                    {
                        token.Subtype = ExcelFormulaTokenSubtype.Logical;
                    }
                    else if (token.Value == "&")
                    {
                        token.Subtype = ExcelFormulaTokenSubtype.Concatenation;
                    }
                    else
                    {
                        token.Subtype = ExcelFormulaTokenSubtype.Math;
                    }

                    _tokens.Add(token);
                    continue;
                }

                if (token.Type == ExcelFormulaTokenType.Operand && token.Subtype == ExcelFormulaTokenSubtype.Nothing)
                {
                    double d;
                    bool isNumber = double.TryParse(token.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out d);
                    if (!isNumber)
                    {
                        if (token.Value == "TRUE" || token.Value == "FALSE")
                        {
                            token.Subtype = ExcelFormulaTokenSubtype.Logical;
                        }
                        else
                        {
                            token.Subtype = ExcelFormulaTokenSubtype.Range;
                        }
                    }
                    else
                    {
                        token.Subtype = ExcelFormulaTokenSubtype.Number;
                    }

                    _tokens.Add(token);
                    continue;
                }

                if (token.Type == ExcelFormulaTokenType.Function)
                {
                    if (token.Value.Length > 0)
                    {
                        if (token.Value.Substring(0, 1) == "@")
                        {
                            token.Value = token.Value.Substring(1);
                        }
                    }
                }

                _tokens.Add(token);
            }
        }
    }
}
