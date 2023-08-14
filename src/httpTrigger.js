const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { notificationApp } = require("./internal/initialize");
const { AppCredential, createMicrosoftGraphClientWithCredential } = require("@microsoft/teamsfx");
const { ResponseType } = require('@microsoft/microsoft-graph-client');

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


    // get the URL value where we can make the call to check if the asynchonous operation to create the Team is complete.
    // this can take a couple of minutes to fully complete.
    const location = team.headers.get('Location');
    // also get the teamId out of the location URL
    const teamId = location.match(/'([^']+)'/)[1];

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
    // const {driveId, generalChannelId} = await getGeneralChannelDriveId(teamId, graphClient);
    const { driveId, generalChannelId } = await getGeneralChannelDriveIdWithRetry(teamId, graphClient);
    console.log(`The driveId of the General channel for team ${teamId} is ${driveId}`);
    console.log(`The id of the General channel for team ${teamId} is ${generalChannelId}`);

    const filePath = 'C:\\Users\\tinsh\\Documents\\Incident.pdf';

    //upload the incident report to the general channel
    const incidentReportUrl = await uploadFileToTeamsChannel(driveId, teamId, generalChannelId, graphClient, filePath);
    console.log(`File uploaded to General channel: ${incidentReportUrl}`);

    for (const target of installations) {
      await target.sendAdaptiveCard(
        AdaptiveCards.declare(notificationTemplate).render({
          title: "New Incident Occurred!",
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

  async function uploadFileToTeamsChannel(driveId, teamId, generalChannelId, graphClient, filePath) {
    const fs = require('fs');
    const path = require('path');
    const fileName = path.basename(filePath);
    const fileSize = fs.statSync(filePath).size;
    const fileContent = fs.createReadStream(filePath);

    //A Microsoft Graph API call to upload a local file to the Teams General channel
  
    try {
      const uploadSession = await graphClient.api(`/drives/${driveId}/root:/General/${fileName}:/createUploadSession`)
        .post({
          item: {
            '@microsoft.graph.conflictBehavior': 'rename',
            name: fileName,
          },
        });
   
      const uploadUrl = uploadSession.uploadUrl;
      const maxChunkSize = 320 * 1024; // 320 KB
      let start = 0;
      let end = maxChunkSize;
      let fileSlice;
     
      while (start < fileSize) {
        if (fileSize - end < 0) {
          end = fileSize;
        }
      
        fileSlice = Buffer.alloc(maxChunkSize);
        const bytesRead = fileContent.read(fileSlice, 0, maxChunkSize, start);
      
        const response = await fetch(uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Range': `bytes ${start}-${start + bytesRead - 1}/${fileSize}`,
          },
          body: fileSlice.slice(0, bytesRead),
        });
      
        start += bytesRead;
        end += maxChunkSize;
      }

      //Complete the upload
      const response = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
          'Content-Length': 0,
        },
      });

      //Get the URL to the uploaded file
      const deepLink = `https://teams.microsoft.com/l/file/${encodeURIComponent(fileName)}/preview?groupId=${teamId}&tenantId=${process.env.TEAMS_APP_TENANT_ID}&channelId=${generalChannelId}`;
      const uploadedFileUrl = `https://teams.microsoft.com/_#/files/tab/${teamId}/${driveId}/${encodeURIComponent(fileName)}`;
      console.log(`File uploaded successfully to ${uploadedFileUrl}`);
      return deepLink;
    } catch (error) {
      console.error(`Error uploading file: ${error}`);
      throw error;
    }
  }

  context.res = {};
};
