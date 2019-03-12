# Teams Provisioning Sample

## Part 1: Solution Overview

Collaborative workspaces are a lot more useful if they're consistent and easy to navigate; this is true of SharePoint, Microsoft Teams, and really any product that has the concept of a workspace. A widely accepted approach is to establish categories of workspaces for use within an enterprise, such as for a department, project, or community of practice. Each category has:

* An initial structure and content
* Standard security and compliance settings
* Well understood ownership and lifetime

Manually creating and configuring each workspace - each Microsoft Team in this case - is labor intensive and sufficiently error-prone that consistency is unlikely in any but the smallest of organizations. Some kind of templating and provisioning system is needed to consistently stamp out, say, Engineering Project Teams or Marketing Campaign Teams.

Microsoft Teams provides two main ways to do this:

1. Teams can be ["cloned"](https://support.office.com/en-us/article/create-a-team-from-an-existing-team-f41a759b-3101-4af6-93bd-6aba0e5d7635), effectively copied with a new name, description, and membership. For each category of Team, a "master" Team is maintained with no content; the "master" Team is then cloned when a new Team is needed. This has the advantage that power users can maintain the "master" Teams without technical assistance.

1. Teams can be created using the Graph API, which provides a very nice [Create team](https://docs.microsoft.com/en-us/graph/api/team-put-teams) call that allows passing in almost every detail of a Team including channels, tabs, and applications. ([This call was in beta](https://docs.microsoft.com/en-us/graph/api/team-post?view=graph-rest-beta) at the time of this writing). This has the advantage that the JSON can be versioned and (at least partly) re-applied to Teams to update them. 

At the time of this writing, there are a number of caveats to cloning; these have been documented (along with work-arounds) by [Laura Kokkarinen](https://laurakokkarinen.com/) in her excellent blog series [Cloning Teams and Configuring Tabs](https://laurakokkarinen.com/cloning-teams-and-configuring-tabs-via-microsoft-graph-prelude/). This project doesn't include the work-arounds, and may be of limited use for cloning until the work-arounds are added or made unnecessary through product improvements.

Instead, the approach of this project - at least initially - is to make it easy to create new Teams via the Graph API, and specifically to do that using a workflow tool such as Microsoft Flow or Azure Logic Apps.

## Teams Templates

Some people call the "master" Team that is meant for cloning a "template." There are also some [built-in templates for education, retail, and healthcare](https://developer.microsoft.com/en-us/office/blogs/deliver-a-consistent-repeatable-microsoft-teams-experience-with-the-launch-of-templates/); these are available out of the box.

I like to think of the Create Team call's [lovely JSON structure](https://docs.microsoft.com/en-us/graph/api/team-post?view=graph-rest-beta#request-2) as a template too, even though the documentation calls it a "team object" or "[team resource type](https://docs.microsoft.com/en-us/graph/api/resources/team?view=graph-rest-beta)." In this solution, this same JSON is stored in files in a SharePoint library. 

With this solution and a set of these JSON templates in place, a Microsoft Flow can easily provision new Teams by adding a message to an Azure queue when its business logic decides it's appropriate. The queue message includes the fields that are different for each team (name, description, owner), and the rest comes from the JSON "template".

## Architecture

Architectural goals were:

* Allow for provisioning at scale, given that Teams creation can be a long-running operation
* Create Teams using an application identity so no user identity is involved (more secure and no added user license required)
* Cleanly separate the details of API calls from the business logic and template content. The core provisioning is managed as a "product" (it can be versioned and deployed in development, staging, production, etc.), and the business rules and templates are managed as "content" (can be edited in place by a business analyst or no-code/low-code developer),

![Solution Architecture](./images/SolutionArchitecture.png)

As shown in the diagram, Team provisioning is initiated through a PowerApp or Flow - or really any program. A Team is provisioned by placing a message on the Create Team Request Queue or Clone Team Request Queue. For example, to request a creation, this JSON is placed in the queue:

~~~JSON
{
  "requestId": "9999",
  "displayName": "My Team",
  "description": "Something new",
  "owner": "someone@mytenant.onmicrosoft.com",
  "jsonTemplate": "EngineeringTemplate"
}
~~~

This will trigger an Azure function which reads the template JSON from SharePoint and merges in the specified displayName, description, and owner, and then creates the Team.

When the Team is created (or has failed for some reason), the Azure Function places a response on the Create Team Completion Queue or Clone Team Completion Queue. This message includes the new Team ID and some other information. Notably, it includes the orginal Request ID, which the two Flows can use to correlate which request they're talking about. For example, the Request ID could contain the ID of an item in a SharePoint list that represents the new Team request; the 2nd Flow can then easily read the SharePoint list item to get all the context it needs to complete its work.

NOTE: At the time of this writing, the Create Team Graph API call is only capable of provisioning a single Team owner if called with an app identity, so this version only handles a single owner. Additional owners may be supported in the future if the underlying API starts supporting them, or by adding them sequentially to the Team. Also note that the current API for adding owners and members is for the underlying O365 Group, and can take up to 2 hours to propagate to the Team, so it might be best to let the initial Owner add other people manually, which is immediate.

