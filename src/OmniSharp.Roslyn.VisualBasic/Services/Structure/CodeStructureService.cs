using System.Collections.Generic;
using System.Collections.Immutable;
using System.Composition;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using Microsoft.CodeAnalysis.Text;
using OmniSharp.Extensions;
using OmniSharp.Mef;
using OmniSharp.Models.V2;
using OmniSharp.Models.V2.CodeStructure;
using OmniSharp.Roslyn.Utilities;
using OmniSharp.Services;

namespace OmniSharp.Roslyn.VisualBasic.Services.Structure
{
    [OmniSharpHandler(OmniSharpEndpoints.V2.CodeStructure, LanguageNames.VisualBasic)]
    public class CodeStructureService : IRequestHandler<CodeStructureRequest, CodeStructureResponse>
    {
        private readonly OmniSharpWorkspace _workspace;
        private readonly IEnumerable<ICodeElementPropertyProvider> _propertyProviders;

        [ImportingConstructor]
        public CodeStructureService(
            OmniSharpWorkspace workspace,
            [ImportMany] IEnumerable<ICodeElementPropertyProvider> propertyProviders)
        {
            _workspace = workspace;
            _propertyProviders = propertyProviders;
        }

        public async Task<CodeStructureResponse> Handle(CodeStructureRequest request)
        {
            // To provide complete code structure for the document wait until all projects are loaded.
            var document = await _workspace.GetDocumentFromFullProjectModelAsync(request.FileName);
            if (document == null)
            {
                return null;
            }

            var elements = await GetCodeElementsAsync(document);

            var response = new CodeStructureResponse
            {
                Elements = elements
            };

            return response;
        }

        private async Task<IReadOnlyList<CodeElement>> GetCodeElementsAsync(Document document)
        {
            var text = await document.GetTextAsync();
            var syntaxRoot = await document.GetSyntaxRootAsync();
            var semanticModel = await document.GetSemanticModelAsync();

            var results = ImmutableList.CreateBuilder<CodeElement>();

            foreach (var node in ((CompilationUnitSyntax)syntaxRoot).Members)
            {
                foreach (var element in CreateCodeElements(node, text, semanticModel))
                {
                    if (element != null)
                    {
                        results.Add(element);
                    }
                }
            }

            return results.ToImmutable();
        }

        private IEnumerable<CodeElement> CreateCodeElements(SyntaxNode node, SourceText text, SemanticModel semanticModel)
        {
            switch (node)
            {
                case TypeBlockSyntax typeDeclaration:
                    yield return CreateCodeElement(typeDeclaration, text, semanticModel);
                    break;
                case DelegateStatementSyntax delegateDeclaration:
                    yield return CreateCodeElement(delegateDeclaration, text, semanticModel);
                    break;
                case EnumBlockSyntax enumDeclaration:
                    yield return CreateCodeElement(enumDeclaration, text, semanticModel);
                    break;
                case NamespaceBlockSyntax namespaceDeclaration:
                    yield return CreateCodeElement(namespaceDeclaration, text, semanticModel);
                    break;
                case MethodBlockBaseSyntax baseMethodDeclaration:
                    yield return CreateCodeElement(baseMethodDeclaration, text, semanticModel);
                    break;
                case PropertyBlockSyntax propertyDeclaration:
                    yield return CreateCodeElement(propertyDeclaration, text, semanticModel);
                    break;
                case FieldDeclarationSyntax fieldDeclaration:
                    foreach (var variableDeclarator in fieldDeclaration.Declarators)
                    {
                        foreach (var name in variableDeclarator.Names)
                        {
                            yield return CreateCodeElement(name, fieldDeclaration, text, semanticModel);
                        }
                    }

                    break;
                case EnumMemberDeclarationSyntax enumMemberDeclarationSyntax:
                    yield return CreateCodeElement(enumMemberDeclarationSyntax, text, semanticModel);
                    break;
            }
        }

        private CodeElement CreateCodeElement(TypeBlockSyntax typeDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(typeDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortTypeFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.TypeFormat)
            };

            AddRanges(builder, typeDeclaration.BlockStatement.AttributeLists.Span, typeDeclaration.Span, typeDeclaration.BlockStatement.Identifier.Span, text);
            AddSymbolProperties(symbol, builder);

            foreach (var member in typeDeclaration.Members)
            {
                foreach (var childElement in CreateCodeElements(member, text, semanticModel))
                {
                    builder.AddChild(childElement);
                }
            }

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(DelegateStatementSyntax delegateDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(delegateDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortTypeFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.TypeFormat),
            };

            AddRanges(builder, delegateDeclaration.AttributeLists.Span, delegateDeclaration.Span, delegateDeclaration.Identifier.Span, text);
            AddSymbolProperties(symbol, builder);

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(EnumBlockSyntax enumDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(enumDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortTypeFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.TypeFormat),
            };

            AddRanges(builder, enumDeclaration.EnumStatement.AttributeLists.Span, enumDeclaration.Span, enumDeclaration.EnumStatement.Identifier.Span, text);
            AddSymbolProperties(symbol, builder);

            foreach (var member in enumDeclaration.Members)
            {
                foreach (var childElement in CreateCodeElements(member, text, semanticModel))
                {
                    builder.AddChild(childElement);
                }
            }

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(NamespaceBlockSyntax namespaceDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(namespaceDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortTypeFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.TypeFormat),
            };

            AddRanges(builder, attributesSpan: default, namespaceDeclaration.Span, namespaceDeclaration.NamespaceStatement.Name.Span, text);

            foreach (var member in namespaceDeclaration.Members)
            {
                foreach (var childElement in CreateCodeElements(member, text, semanticModel))
                {
                    builder.AddChild(childElement);
                }
            }

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(MethodBlockBaseSyntax baseMethodDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(baseMethodDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortMemberFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.MemberFormat),
            };

            AddRanges(builder, baseMethodDeclaration.BlockStatement.AttributeLists.Span, baseMethodDeclaration.Span, GetNameSpan(baseMethodDeclaration), text);
            AddSymbolProperties(symbol, builder);

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(PropertyBlockSyntax propertyDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(propertyDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortMemberFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.MemberFormat),
            };

            AddRanges(builder, propertyDeclaration.PropertyStatement.AttributeLists.Span, propertyDeclaration.Span, propertyDeclaration.PropertyStatement.Identifier.Span, text);
            AddSymbolProperties(symbol, builder);

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(ModifiedIdentifierSyntax modifiedIdentifier, FieldDeclarationSyntax fieldDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(modifiedIdentifier);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortMemberFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.MemberFormat),
            };

            AddRanges(builder, fieldDeclaration.AttributeLists.Span, modifiedIdentifier.Parent.Span, modifiedIdentifier.Identifier.Span, text);
            AddSymbolProperties(symbol, builder);

            return builder.ToCodeElement();
        }

        private CodeElement CreateCodeElement(EnumMemberDeclarationSyntax enumMemberDeclaration, SourceText text, SemanticModel semanticModel)
        {
            var symbol = semanticModel.GetDeclaredSymbol(enumMemberDeclaration);
            if (symbol == null)
            {
                return null;
            }

            var builder = new CodeElement.Builder
            {
                Kind = symbol.GetKindString(),
                Name = symbol.ToDisplayString(SymbolDisplayFormats.ShortMemberFormat),
                DisplayName = symbol.ToDisplayString(SymbolDisplayFormats.MemberFormat),
            };

            AddRanges(builder, enumMemberDeclaration.AttributeLists.Span, enumMemberDeclaration.Span, enumMemberDeclaration.Identifier.Span, text);
            AddSymbolProperties(symbol, builder);

            return builder.ToCodeElement();
        }

        private static TextSpan GetNameSpan(MethodBlockBaseSyntax baseMethodDeclaration)
        {
            switch (baseMethodDeclaration)
            {
                case AccessorBlockSyntax accessorBlock:
                    return accessorBlock.AccessorStatement.AccessorKeyword.Span;
                case ConstructorBlockSyntax constructorBlock:
                    return constructorBlock.SubNewStatement.NewKeyword.Span;
                case MethodBlockSyntax methodBlock:
                    return methodBlock.SubOrFunctionStatement.Identifier.Span;
                case OperatorBlockSyntax operatorBlock:
                    return operatorBlock.OperatorStatement.OperatorToken.Span;
                default:
                    return default;
            }
        }

        private static void AddRanges(CodeElement.Builder builder, TextSpan attributesSpan, TextSpan fullSpan, TextSpan nameSpan, SourceText text)
        {
            if (attributesSpan != default)
            {
                builder.AddRange(SymbolRangeNames.Attributes, text.GetRangeFromSpan(attributesSpan));
            }

            if (fullSpan != default)
            {
                builder.AddRange(SymbolRangeNames.Full, text.GetRangeFromSpan(fullSpan));
            }

            if (nameSpan != default)
            {
                builder.AddRange(SymbolRangeNames.Name, text.GetRangeFromSpan(nameSpan));
            }
        }

        private void AddSymbolProperties(ISymbol symbol, CodeElement.Builder builder)
        {
            var accessibility = symbol.GetAccessibilityString();
            if (accessibility != null)
            {
                builder.AddProperty(SymbolPropertyNames.Accessibility, accessibility);
            }

            builder.AddProperty(SymbolPropertyNames.Static, symbol.IsStatic);

            foreach (var propertyProvider in _propertyProviders)
            {
                foreach (var (name, value) in propertyProvider.ProvideProperties(symbol))
                {
                    builder.AddProperty(name, value);
                }
            }
        }
    }
}
