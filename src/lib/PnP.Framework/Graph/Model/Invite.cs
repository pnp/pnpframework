namespace PnP.Framework.Graph.Model
{
    public class Invite
    {
        public string InvitedUserEmailAddress { get; set; }
        public string InvitedUserDisplayName { get; set; }
        public string InviteRedirectUrl { get; set; }
        public bool SendInvitationMessage { get; set; }
        public InvitedUserMessageInfo InvitedUserMessageInfo { get; set; }
    }
}
