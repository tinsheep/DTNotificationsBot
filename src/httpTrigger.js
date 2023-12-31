const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { notificationApp } = require("./internal/initialize");
const { AppCredential, createMicrosoftGraphClientWithCredential } = require("@microsoft/teamsfx");
const { ResponseType } = require('@microsoft/microsoft-graph-client');
const { uploadToTeamsGeneralChannel } = require('./uploadToTeamsGeneralChannel');

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
module.exports = async function (context, req) {
  const pageSize = 100;
  let continuationToken = undefined;
  //Loop through all the installations and send notification to each target. I dont know if we need to do this for each target becasue we dont want to create multiple instantes of a Team but leave the logic for now.
  do {
    const pagedData = await notificationApp.notification.getPagedInstallations(
      pageSize,
      continuationToken
    );
    const installations = pagedData.data;
    continuationToken = pagedData.continuationToken;


    const appAuthConfig = {
      authorityHost: process.env.M365_AUTHORITY_HOST,
      clientId: process.env.BOT_ID,
      tenantId: process.env.TEAMS_APP_TENANT_ID,
      clientSecret: process.env.BOT_PASSWORD
    }
    const appCredential = new AppCredential(appAuthConfig)
    const graphClient = createMicrosoftGraphClientWithCredential(appCredential);

    //create a new team based on the incomming request

    let teamId;
    let driveId;
    let generalChannelId;
    const requestBody = req.body;

    //select the email from the request body of the first owner role in the members array
    const ownerMember = requestBody.members.find(member => member.role && member.role.includes('owner'));
    if (!ownerMember) {
      throw new Error('No owner role defined in the request');
    }
    const ownerEmail = ownerMember.email;


    //debug switch to create a new team or use an existing team to save time debugging
    if (process.env.CREATE_TEAM === 'true') {

      //create Incident response team from a teamsTemplate
      const teamTemplate = {
        'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'' + requestBody.templateId  +  '\')',
        displayName: requestBody.incidentName,
        description: requestBody.incidentDescription,
        members:[
            {
               '@odata.type': '#microsoft.graph.aadUserConversationMember',
               roles:[
                  'owner'
               ],
              'user@odata.bind': 'https://graph.microsoft.com/v1.0/users(\'' + ownerEmail + '\')'
            }
        ]
      }

      const team = await graphClient
          .api('/teams')
          .responseType(ResponseType.RAW)
          .post(teamTemplate);
      console.log("Client Request ID: ", team.headers.get('client-request-id'));


      // get the URL value where we can make the call to check if the asynchronous operation to create the Team is complete.
      // this can take a couple of minutes to fully complete.
      const location = team.headers.get('Location');
      // also get the teamId out of the location URL
      teamId = location.match(/'([^']+)'/)[1];

      let teamStatus = "inProgress";
      let waitingMessage = "Waiting for site creation to complete";
      while (teamStatus != "succeeded") {
        const checkStatusResponse = await graphClient.api(location).get();
        teamStatus = checkStatusResponse.status;
        console.log(`${waitingMessage}.`);
        waitingMessage += ".";
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }
      console.log("Site creation completed successfully!");

      // I need the driveId so i can post the incident report to the channel
      const { driveId: tempDriveId, generalChannelId: tempGeneralChannelId } = await getGeneralChannelDriveIdWithRetry(teamId, graphClient);
      driveId = tempDriveId;
      generalChannelId = tempGeneralChannelId;
      console.log(`The driveId of the General channel for team ${teamId} is ${driveId}`);
      console.log(`The id of the General channel for team ${teamId} is ${generalChannelId}`);
    } else {
      //reuse an existing team
      teamId = process.env.TEAM_ID;
      driveId = process.env.DRIVE_ID;
      generalChannelId = process.env.GENERAL_CHANNEL_ID;
    }

    const filePath = 'C:\\Users\\tinsh\\Documents\\Incident.pdf';

    //upload the incident report to the general channel
    //I need to also pass in the teamId and the generalChannelId to properly construct the deep link to the file
    const incidentReportUrl = await  uploadToTeamsGeneralChannel(driveId, filePath, graphClient, teamId, generalChannelId);
 
    console.log(`File uploaded to General channel: ${incidentReportUrl}`);

    //send a notification to the installed app
    for (const target of installations) {
      await target.sendAdaptiveCard(
        AdaptiveCards.declare(notificationTemplate).render({
          title: "New Incident Workspace Created",
          appName: "Disaster Tech",
          description: `A new incident workspace was created. Click the button below to view the incident report:`,
          notificationUrl: incidentReportUrl,
        })
      );


      //now that the incident report is uploaded to the general channel, we can add additional members to the team
      await addMembersToTeam(teamId, requestBody.members, graphClient);


      /****** To distinguish different target types ******/
      /** "Channel" means this bot is installed to a Team (default to notify General channel)
      if (target.type === NotificationTargetType.Channel) {
        // Directly notify the Team (to the default General channel)
        await target.sendAdaptiveCard(...);
        // List all channels in the Team then notify each channel
        const channels = await target.channels();
        for (const channel of channels) {
          await channel.sendAdaptiveCard(...);
        }
        // List all members in the Team then notify each member
        const pageSize = 100;
        let continuationToken = undefined;
        do {
          const pagedData = await target.getPagedMembers(pageSize, continuationToken);
          const members = pagedData.data;
          continuationToken = pagedData.continuationToken;

          for (const member of members) {
            await member.sendAdaptiveCard(...);
          }
        } while (continuationToken);
      }
      **/

      /** "Group" means this bot is installed to a Group Chat
      if (target.type === NotificationTargetType.Group) {
        // Directly notify the Group Chat
        await target.sendAdaptiveCard(...);
        // List all members in the Group Chat then notify each member
        const pageSize = 100;
        let continuationToken: string | undefined = undefined;
        do {
          const pagedData = await target.getPagedMembers(pageSize, continuationToken);
          const members = pagedData.data;
          continuationToken = pagedData.continuationToken;

          for (const member of members) {
            await member.sendAdaptiveCard(...);
          }
        } while (continuationToken);
      }
      **/

      /** "Person" means this bot is installed as a Personal app
      if (target.type === NotificationTargetType.Person) {
        // Directly notify the individual person
        await target.sendAdaptiveCard(...);
      }
      **/
    }
  } while (continuationToken);

  /** You can also find someone and notify the individual person
  const member = await notificationApp.notification.findMember(
    async (m) => m.account.email === "someone@contoso.com"
  );
  await member?.sendAdaptiveCard(...);
  **/

  /** Or find multiple people and notify them
  const members = await notificationApp.notification.findAllMembers(
    async (m) => m.account.email?.startsWith("test")
  );
  for (const member of members) {
    await member.sendAdaptiveCard(...);
  }
  **/

  // Supporting functions

  function generateGUID() {
    let guid = "";
    for (let i = 0; i < 32; i++) {
      guid += Math.floor(Math.random() * 16).toString(16);
    }
    return guid;
  }

  async function addMembersToTeam(teamId, members, graphClient) {
    for (const member of members) {
      let user;
      try {
        user = await graphClient.api(`/users/${member.email}`).get();
      } catch (error) {
        console.log(`Error getting user with email ${member.email}: ${error.message}`);
        // Retry up to 3 times with a delay of 2 seconds between retries
        // This is to handle the case where the user is not found immediately after creation
        // This can happen if the user is created in Azure AD but not yet replicated to Microsoft Graph
        // it might also be the case that the user is not found because it is a guest user
        // that does not exist in the tenant.
        for (let i = 0; i < 3; i++) {
          console.log(`Retrying... Attempt ${i + 1}`);
          await new Promise(resolve => setTimeout(resolve, 2000));
          try {
            user = await graphClient.api(`/users/${member.email}`).get();
            break;
          } catch (error) {
            console.log(`Error getting user with email ${member.email}: ${error.message}`);
          }
        }
      }

      // Check if the user is already a member of the team
      const teamMembers = await graphClient.api(`/teams/${teamId}/members`).get();
      const existingMember = teamMembers.value.find((m) => m.email && m.email.toLowerCase() === user?.userPrincipalName?.toLowerCase());
      if (existingMember) {
        console.log(`User ${user.userPrincipalName} is already a member of the team.`);
        continue;
      }

      // Determine the role of the member
      let role;
      if (member.role && member.role.includes('owner')) {
        role = 'owner';
      } else if (member.role && member.role.includes('member')) {
        role = []; // Set role to an empty array as this means the user is a member.
      } else if (member.role && member.role.includes('guest')) {
        let guestUser;
        try {
          guestUser = await graphClient.api(`/users?$filter=startswith(mail, '${encodeURIComponent(member.email)}')&$select=userType,id,userPrincipalName`).get();
          guestUser.id = guestUser.value[0].id;
          console.log(`User ${member.email} is already a guest in the tenant.`);
        } catch (error) {
          if (error.statusCode === 404) {
            console.log(`User ${member.email} is not found in the tenant. Inviting...`);
            // Send an invite to the guest to join the tenant
            const invite = {
              invitedUserEmailAddress: member.email,
              inviteRedirectUrl: 'https://teams.microsoft.com',
              sendInvitationMessage: true,
              roles: ['guest']
            };
            const invitation = await graphClient.api(`/invitations`).post(invite);
            console.log(`Sent invitation to ${member.email} to join the tenant as a guest.`);

            // Update the user object with the invitation ID
            guestUser.id = invitation.invitedUser.id;
          } else {
            console.log(`Error checking if user ${member.email} is a guest in the tenant: ${error.message}`);
          }
        }

        // Add the guest to the team
        // Note: This call is not supported with application permissions and will result in a 403 Forbidden error.
        // You must use delegated permissions to add a guest to a team.
        // uncomment when you have the correct permissions.
        // const guestToAdd = {
        //   '@odata.type': '#microsoft.graph.aadUserConversationMember',
        //   'roles': ['guest'],
        //   'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${guestUser.id}`
        // };
        // await graphClient.api(`/teams/${teamId}/members`).post(guestToAdd);
        // console.log(`Added guest ${member.email} to the team as a guest.`);
        continue;
      } else {
        console.log(`Unknown role for user ${member.email}. Skipping.`);
        continue;
      }

      // Add the member to the team
      const memberToAdd = {
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles: role,
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users/${user.id}`
      };
      await graphClient.api(`/teams/${teamId}/members`).post(memberToAdd);
      console.log(`Added user ${user.userPrincipalName} to the team as a ${role.length ? role : 'member'}.`);
    }
  }

  async function getGeneralChannelDriveIdWithRetry(teamId, graphClient, maxRetries = 5, retryDelay = 2000) {
    let retries = 0;
    while (true) {
      try {
        const { driveId, generalChannelId } = await getGeneralChannelDriveId(teamId, graphClient);
        return { driveId, generalChannelId };
      } catch (error) {
        if (error.message === 'Folder location for this channel is not ready yet, please try again later.' && retries < maxRetries) {
          retries++;
          console.log(`Retrying... Attempt ${retries}`);
          await new Promise(resolve => setTimeout(resolve, retryDelay));
        } else {
          throw error;
        }
      }
    }
  }

  async function getGeneralChannelDriveId(teamId, client) {
    const channels = await client.api(`/teams/${teamId}/channels`).filter(`displayName eq 'General'`).select('id').get();
    const generalChannelId = channels.value[0].id;
    const drive = await client.api(`/teams/${teamId}/channels/${generalChannelId}/filesFolder`).get();
    const driveId = drive.parentReference.driveId;
    return {driveId, generalChannelId};
  }

  context.res = {};
};
