import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  TaskModuleRequest,
  TaskModuleResponse,
} from "botbuilder";

export class ActionApp extends TeamsActivityHandler {
  //Action
  public override async handleTeamsMessagingExtensionSubmitAction(
    context: TurnContext,
    action: MessagingExtensionAction
  ): Promise<MessagingExtensionActionResponse> {
    // The user has chosen to create a card by choosing the 'Create Card' context menu command.
    console.log(`HANDLE SUBMIT ACTION, action: ${JSON.stringify(action)}`);

    const commandId: string = action.commandId;
    const commandData = action.data;
    
    switch (commandId) {
      case "createCard":
        {
          const attachment = CardFactory.adaptiveCard({
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            type: "AdaptiveCard",
            version: "1.4",
            body: [
              {
                type: "TextBlock",
                text: `${commandData.title}`,
                wrap: true,
                size: "Large",
              },
              {
                type: "TextBlock",
                text: `${commandData.subTitle}`,
                wrap: true,
                size: "Medium",
              },
              {
                type: "TextBlock",
                text: `${commandData.text}`,
                wrap: true,
                size: "Small",
              },
            ],
          });
          return {
            composeExtension: {
              type: "result",
              attachmentLayout: "list",
              attachments: [attachment],
            },
          };
        }
      
      default:
        return;
    }  
  }

  public override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`TASK MODULE FETCH. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return;
  }
}
