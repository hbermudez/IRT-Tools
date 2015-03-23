# IRT-Tools
Tools for Incident Response Teams

This project was born to deal with the increasing number of users sending us suspicious emails for analysis (containing either malicious links or attachments or both ). The task became more arduous and painful since we were using Proofpoint, so we needed to manually disentangle all the stuff added to URLs by Proofpoint before even start the analysis.

So I first created created a powershell script to help me out on this and later to an Outlook AddIn.

The code is lacking of some comments and probably this functionality could be achieved in a better and mode efficient way but I am claiming that this is a finished effort, on the contrary, I would like to receive feedbacks and/or contributions from the community so we could come up with something more robust and comprehensive.

Current Features:
- Ability to extract emails' header.
- Ability to extract emails's attachments and append a safeguard extension so you could manipulate/transport the attachment in a safely manner (attachment.pdf.quarantine)
- Ability to disentangle all the Proofpoint appends and recover a URL that can analyzed by OSINTs and/or other means. However it also attach a safeguard (Ex. hxxp://potentiallymaliciousurl.com) 
- Create a report per email containing the emails header, email's sender,  all the IPs on the email and also the URLs. 
- Ability to check every IP and every URL found on the email against VirusTotal (other OSINT engines can be added). When you check multiple emails at once, this feature will create a report of which email is worth it to look first based on the number of alerts it generated.

I have not figure it out how to add screenshots to github, so go to my webpage, http://hlsbinarysolutions.com/blog to find a more descriptive how-to.
