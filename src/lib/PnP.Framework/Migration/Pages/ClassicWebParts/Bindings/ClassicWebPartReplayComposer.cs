using PnP.Framework.Migration.Lists.Planning;
using PnP.Framework.Migration.Pages.Content;
using System;
using System.Collections.Generic;
using System.IO;

namespace PnP.Framework.Migration.Pages.ClassicWebParts.Bindings
{
    internal static class ClassicWebPartReplayComposer
    {
        public static string Compose(
            ClassicWebPartSnapshot captured,
            ClassicWebPartAction action,
            ClassicListWebPartBindingSnapshot binding,
            ListMaterializationReceipt listReceipt,
            string targetPageServerRelativeUrl,
            IEnumerable<PageTextReplacement> replacements)
        {
            if (captured == null)
            {
                throw new ArgumentNullException(nameof(captured));
            }
            if (action == null || action.SourceWebPartId != captured.Id)
            {
                throw new InvalidDataException("The Web Part action does not match the captured Web Part.");
            }
            if (action.Disposition == ClassicWebPartDisposition.CopyCaptured)
            {
                return PageTextTransformer.Rewrite(captured.ExportXml, replacements);
            }
            if (action.Disposition != ClassicWebPartDisposition.RebindListAfterMaterialization)
            {
                throw new InvalidOperationException("A blocked Web Part action cannot be replayed.");
            }
            if (binding == null || binding.SourceWebPartId != captured.Id)
            {
                throw new InvalidDataException("The list-bound Web Part has no matching captured binding.");
            }
            if (listReceipt == null
                || listReceipt.SourceWebId != binding.SourceListWebId
                || listReceipt.SourceListId != binding.SourceListId)
            {
                throw new InvalidDataException("The list-bound Web Part has no matching List materialization receipt.");
            }
            if (!binding.SourceViewId.HasValue)
            {
                throw new InvalidDataException("The list-bound Web Part has no captured source View identity.");
            }

            Guid targetViewId;
            if (!listReceipt.TargetViewIds.TryGetValue(binding.SourceViewId.Value, out targetViewId))
            {
                throw new InvalidDataException("The captured source View has no target View identity in the List receipt.");
            }

            var rewritten = ClassicListWebPartRewriter.Rewrite(binding, new ClassicListWebPartTargetMap
            {
                SourceWebId = binding.SourceListWebId,
                SourceListId = binding.SourceListId,
                SourceViewId = binding.SourceViewId,
                TargetWebId = listReceipt.TargetWebId,
                TargetListId = listReceipt.TargetListId,
                TargetViewId = targetViewId,
                TargetListServerRelativeUrl = listReceipt.TargetRootFolderServerRelativeUrl,
                TargetPageServerRelativeUrl = targetPageServerRelativeUrl,
                RenderingResourceRewrites = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            });
            return PageTextTransformer.Rewrite(rewritten.ExportXml, replacements);
        }
    }
}
