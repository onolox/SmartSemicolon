using System;
using System.ComponentModel.Composition;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;

namespace SmartSemicolon
{
    /// <summary> Ponto e vírgula coloca um ponto e vírgula no fim da linha, não importando onde se esteja na linha,
    /// um backspace imediatamente faz com que o ; seja colocado na posição inicial do input.</summary>
    internal sealed class VirgulaAdornment
    {
        [Import]
        private DTE2 applicationObject;
        private TextDocumentKeyPressEvents textDocKeyEvents;
        private Boolean pontoVirgula = false;
        private int colunaPontoVirgula, linhaPontoVirgula, colunaAtual;

        /// <summary>Initializes a new instance of the <see cref="VirgulaAdornment"/> class.</summary>
        /// <param name="view">Text view to create the adornment for</param>
        public VirgulaAdornment(IWpfTextView view, SVsServiceProvider sp)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }
            applicationObject = sp.GetService(typeof(DTE)) as DTE2;

            Events2 eventosTeclado = (Events2)applicationObject.Events;
            textDocKeyEvents = eventosTeclado.get_TextDocumentKeyPressEvents(null);

            // Connect to the BeforeKeyPress delegate exposed from the TextDocumentKeyPressEvents object retrieved above.
            try
            { 
                textDocKeyEvents.BeforeKeyPress += new _dispTextDocumentKeyPressEvents_BeforeKeyPressEventHandler(BeforeKeyPress);
            }
            catch (Exception) { }
        }

        private void BeforeKeyPress(string Keypress, TextSelection Selection, bool InStatementCompletion, ref bool CancelKeypress)
        {
            if (Keypress == ";")
            {
                String nomeArquivo = Selection.Parent.Parent.Name;
                if (nomeArquivo.EndsWith(".cs") || nomeArquivo.EndsWith(".js"))
                {
                    bool isEndOfLine = false;
                    if (linhaPontoVirgula != Selection.ActivePoint.Line)
                    { // Checa se a linha atual da inserção é a mesma, caso não seja cancela todas ações
                        pontoVirgula = false;
                        linhaPontoVirgula = Selection.ActivePoint.Line;
                    }
                    if (Selection.ActivePoint.LineCharOffset != colunaAtual)
                    { // Checa se é a mesma coluna
                        pontoVirgula = false;
                        colunaAtual = Selection.ActivePoint.LineCharOffset;
                        isEndOfLine = Selection.ActivePoint.AtEndOfLine;
                    }

                    CancelKeypress = true;
                    colunaPontoVirgula = Selection.ActivePoint.LineCharOffset; // Salva posição de input na linha caso necessário o backspace

                    applicationObject.UndoContext.Open("ponto virgula", false);

                    Selection.EndOfLine(false);
                    Selection.CharLeft(true);
                    if (Selection.Text != ";")                  // Se não existir ; no fim da linha insere ; no fim, mas | fica na mesma posição

                    {
                        Selection.CharRight(false);
                        Selection.Insert(";");
                        if (!isEndOfLine)
                        {
                            Selection.MoveToLineAndOffset(linhaPontoVirgula, colunaPontoVirgula);
                        }

                        pontoVirgula = true;
                    }
                    else if (!Selection.ActivePoint.AtEndOfLine)
                    {
                        Selection.CharRight(false);
                        Selection.MoveToLineAndOffset(linhaPontoVirgula, colunaPontoVirgula);
                        Selection.Insert(";");
                        pontoVirgula = false;
                    }

                    applicationObject.UndoContext.Close();
                    colunaAtual = Selection.ActivePoint.LineCharOffset;
                    isEndOfLine = Selection.ActivePoint.AtEndOfLine;
                }
            }
            else if (Keypress == "\b") // Backspace  
            {
                if (pontoVirgula)                // Se clicar backspace o ponto e virgula volta a posição inicial do carret
                {
                    if ((Selection.ActivePoint.Line == linhaPontoVirgula) && (Selection.ActivePoint.LineCharOffset == colunaAtual))
                    {
                        CancelKeypress = true;
                        applicationObject.ActiveDocument.Undo();
                        applicationObject.UndoContext.Open("Virgula", false);

                        //Selection.ActivePoint.CreateEditPoint().Delete(-1);
                        Selection.MoveToLineAndOffset(linhaPontoVirgula, colunaPontoVirgula, false);
                        Selection.Insert(";", (int)vsInsertFlags.vsInsertFlagsCollapseToEnd);


                        applicationObject.UndoContext.Close();
                    }
                }
                pontoVirgula = false;
            }
            else pontoVirgula = false;
        }
    }
}