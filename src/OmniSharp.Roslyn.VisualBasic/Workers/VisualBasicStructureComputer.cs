using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using OmniSharp.Abstractions.Services;
using OmniSharp.Models;
using OmniSharp.Models.MembersTree;

namespace OmniSharp
{
    public class VisualBasicStructureComputer : VisualBasicSyntaxWalker
    {
        private readonly Stack<FileMemberElement> _roots = new Stack<FileMemberElement>();
        private readonly IEnumerable<ISyntaxFeaturesDiscover> _featureDiscovers;
        private string _currentProject;
        private SemanticModel _semanticModel;

        public static async Task<IEnumerable<FileMemberElement>> Compute(
            IEnumerable<Document> documents,
            IEnumerable<ISyntaxFeaturesDiscover> featureDiscovers)
        {
            var root = new FileMemberElement() { ChildNodes = new List<FileMemberElement>() };
            var visitor = new VisualBasicStructureComputer(root, featureDiscovers);

            foreach (var document in documents)
            {
                await visitor.Process(document);
            }

            return root.ChildNodes;
        }

        public static Task<IEnumerable<FileMemberElement>> Compute(IEnumerable<Document> documents) 
            => Compute(documents, Enumerable.Empty<ISyntaxFeaturesDiscover>());

        private VisualBasicStructureComputer(FileMemberElement root, IEnumerable<ISyntaxFeaturesDiscover> featureDiscovers)
        {
            _featureDiscovers = featureDiscovers ?? Enumerable.Empty<ISyntaxFeaturesDiscover>();
            _roots.Push(root);
        }

        private async Task Process(Document document)
        {
            _currentProject = document.Project.Name;

            if (_featureDiscovers.Any(dis => dis.NeedSemanticModel))
            {
                _semanticModel = await document.GetSemanticModelAsync();
            }

            var syntaxRoot = await document.GetSyntaxRootAsync();
            (syntaxRoot as VisualBasicSyntaxNode)?.Accept(this);
        }

        private FileMemberElement AsNode(SyntaxNode node, string text, Location location, TextSpan attributeSpan)
        {
            var ret = new FileMemberElement();
            var lineSpan = location.GetLineSpan();
            ret.Projects = new List<string>();
            ret.ChildNodes = new List<FileMemberElement>();
            ret.Kind = node.Kind().ToString();
            ret.AttributeSpanStart = attributeSpan.Start;
            ret.AttributeSpanEnd = attributeSpan.End;
            ret.Location = new QuickFix()
            {
                Text = text,
                FileName = lineSpan.Path,
                Line = lineSpan.StartLinePosition.Line,
                Column = lineSpan.StartLinePosition.Character,
                EndLine = lineSpan.EndLinePosition.Line,
                EndColumn = lineSpan.EndLinePosition.Character
            };

            foreach (var featureDiscover in _featureDiscovers)
            {
                var features = featureDiscover.Discover(node, _semanticModel);
                foreach (var feature in features)
                {
                    ret.Features.Add(feature);
                }
            }

            return ret;
        }

        private FileMemberElement AsChild(SyntaxNode node, string text, Location location, TextSpan attributeSpan)
        {
            var child = AsNode(node, text, location, attributeSpan);
            var childNodes = ((List<FileMemberElement>)_roots.Peek().ChildNodes);

            // Prevent inserting the same node multiple times
            // but make sure to insert them at the right spot
            var idx = childNodes.BinarySearch(child);
            if (idx < 0)
            {
                ((List<string>)child.Projects).Add(_currentProject);
                childNodes.Insert(~idx, child);
                return child;
            }
            else
            {
                ((List<string>)childNodes[idx].Projects).Add(_currentProject);
                return childNodes[idx];
            }
        }

        private FileMemberElement AsParent(SyntaxNode node, string text, Action fn, Location location, TextSpan attributeSpan)
        {
            var child = AsChild(node, text, location, attributeSpan);
            _roots.Push(child);
            fn();
            _roots.Pop();
            return child;
        }

        public override void VisitClassBlock(ClassBlockSyntax node)
        {
            AsParent(node, node.ClassStatement.Identifier.Text, () => base.VisitClassBlock(node), node.ClassStatement.Identifier.GetLocation(), node.ClassStatement.AttributeLists.Span);
        }

        public override void VisitInterfaceBlock(InterfaceBlockSyntax node)
        {
            AsParent(node, node.InterfaceStatement.Identifier.Text, () => base.VisitInterfaceBlock(node), node.InterfaceStatement.Identifier.GetLocation(), node.BlockStatement.AttributeLists.Span);
        }

        public override void VisitEnumBlock(EnumBlockSyntax node)
        {
            AsParent(node, node.EnumStatement.Identifier.Text, () => base.VisitEnumBlock(node), node.EnumStatement.Identifier.GetLocation(), node.EnumStatement.AttributeLists.Span);
        }

        public override void VisitEnumMemberDeclaration(EnumMemberDeclarationSyntax node)
        {
            AsChild(node, node.Identifier.Text, node.Identifier.GetLocation(), node.AttributeLists.Span);
        }

        public override void VisitPropertyBlock(PropertyBlockSyntax node)
        {
            AsChild(node, node.PropertyStatement.Identifier.Text, node.PropertyStatement.Identifier.GetLocation(), node.PropertyStatement.AttributeLists.Span);
        }

        public override void VisitMethodBlock(MethodBlockSyntax node)
        {
            AsChild(node, node.SubOrFunctionStatement.Identifier.Text, node.SubOrFunctionStatement.Identifier.GetLocation(), node.SubOrFunctionStatement.AttributeLists.Span);
        }

        public override void VisitFieldDeclaration(FieldDeclarationSyntax node)
        {
            AsChild(node, node.Declarators.First().Names.First().Identifier.Text, node.Declarators.First().Names.First().Identifier.GetLocation(), node.AttributeLists.Span);
        }

        public override void VisitEventBlock(EventBlockSyntax node)
        {
            AsChild(node, node.EventStatement.Identifier.Text, node.EventStatement.Identifier.GetLocation(), node.EventStatement.AttributeLists.Span);
        }

        public override void VisitStructureBlock(StructureBlockSyntax node)
        {
            AsParent(node, node.StructureStatement.Identifier.Text, () => base.VisitStructureBlock(node), node.StructureStatement.Identifier.GetLocation(), node.StructureStatement.AttributeLists.Span);
        }

        public override void VisitConstructorBlock(ConstructorBlockSyntax node)
        {
            AsChild(node, node.SubNewStatement.NewKeyword.Text, node.SubNewStatement.NewKeyword.GetLocation(), node.SubNewStatement.AttributeLists.Span);
        }
    }
}
