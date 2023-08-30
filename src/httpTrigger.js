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

    let teamId;
    let driveId;
    let generalChannelId;
    //debug switch to create a new team or use an existing team to save time debugging
    if (process.env.CREATE_TEAM === 'true') {

      //create Incident response team from a teamsTemplate
      const teamTemplate = {
        'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'' + process.env.TEAMS_TEMPLATE_ID  +  '\')',
        displayName: process.env.INCIDENT_NAME,
        description: process.env.INCIDENT_DESCRIPTION,
        members:[
            {
               '@odata.type': '#microsoft.graph.aadUserConversationMember',
               roles:[
                  'owner'
               ],
               'user@odata.bind': 'https://graph.microsoft.com/v1.0/users/' + process.env.MOD_ID
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
    const incidentReportUrl = await  uploadToTeamsGeneralChannel(driveId, filePath, graphClient);
 
    console.log(`File uploaded to General channel: ${incidentReportUrl}`);

    for (const target of installations) {
      await target.sendAdaptiveCard(
        AdaptiveCards.declare(notificationTemplate).render({
          title: "New incident workspace created.",
          appName: "Disaster Tech",
          description: `Welcome to the new incident team. Here is the Incident Action Plan:  ${target.type}`,
          notificationUrl: incidentReportUrl,
        })
      );


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
