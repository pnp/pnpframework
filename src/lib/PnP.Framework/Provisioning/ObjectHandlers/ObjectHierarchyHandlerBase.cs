using Microsoft.Online.SharePoint.TenantAdministration;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.Model.Configuration;
using System.Threading.Tasks;

namespace PnP.Framework.Provisioning.ObjectHandlers
{
    internal delegate bool ShouldProvisionSequenceTest(Tenant web, Model.ProvisioningHierarchy hierarchy);

    internal abstract class ObjectHierarchyHandlerBase
    {
        internal bool? _willExtract;
        internal bool? _willProvision;

        private bool _reportProgress = true;
        public abstract string Name { get; }

        public bool ReportProgress
        {
            get { return _reportProgress; }
            set { _reportProgress = value; }
        }

        public ProvisioningMessagesDelegate MessagesDelegate { get; set; }

        public abstract bool WillProvision(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ApplyConfiguration configuration);

        public abstract bool WillExtract(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, ExtractConfiguration configuration);

        public abstract Task<TokenParser> ProvisionObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, string sequenceId, TokenParser parser, ApplyConfiguration configuration);

        public abstract Task<ProvisioningHierarchy> ExtractObjects(Tenant tenant, Model.ProvisioningHierarchy hierarchy, ExtractConfiguration configuration);

        internal void WriteMessage(string message, ProvisioningMessageType messageType)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate(message, messageType);
            }
        }

        internal void WriteSubProgress(string title, string message, int step, int total)
        {
            if (MessagesDelegate != null)
            {
                MessagesDelegate($"{title}|{message}|{step}|{total}", ProvisioningMessageType.Progress);
            }
        }
    }
}
