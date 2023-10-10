import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  TaskModuleRequest,
  TaskModuleResponse,
} from "botbuilder";

const urlDialogTriggerValue = "requestUrl";
const cardDialogTriggerValue = "requestCard";
const messagePageTriggerValue = "requestMessage";
const noResponseTriggerValue = "requestNoResponse";

const adaptiveCardBotJson = {
  "contentType": "application/vnd.microsoft.card.adaptive",
  "content": {
      "type": "AdaptiveCard",
      "body": [
          {
              "type": "TextBlock",
              "text": "Here is a ninja cat:"
          },
          {
              "type": "Image",
              "url": "http://adaptivecards.io/content/cats/1.png",
              "size": "Medium"
          }
      ],
      "actions": [
          {
              "data": { data: urlDialogTriggerValue },
              "type": "Action.Submit",
              "title": "Request URL Dialog"
          },
          {
            "data": { data: cardDialogTriggerValue },
            "type": "Action.Submit",
            "title": "Request Card Dialog"
          },
          {
            "data": { data: messagePageTriggerValue },
            "type": "Action.Submit",
            "title": "Request Message"
          },
          {
            "data": { data: noResponseTriggerValue },
            "type": "Action.Submit",
            "title": "Request No Response (close Dialog)"
          },
      ],
      "version": "1.0"
  }
}

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

    if (commandData.data === urlDialogTriggerValue) {
      return this.createUrlTaskModuleMEResponse();
    } else if (commandData.data === cardDialogTriggerValue) {
      return this.createCardTaskModuleMEResponse();
    }
    
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
            actions: [
              {
                type: "Action.Submit",
                title: "Show URL Task Module",
                value: {
                  type: 'task/fetch',
                  data: urlDialogTriggerValue
                }
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

  private getRandomIntegerBetween(min: number, max: number): number {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  private createUrlTaskModuleMEResponse(): Promise<MessagingExtensionActionResponse> {
    return Promise.resolve({
      task: {
        type: 'continue',
        value: {
          url: `https://helloworld36cffe.z5.web.core.windows.net/index.html?randomNumber=${this.getRandomIntegerBetween(1, 1000)}#/tab`,
          fallbackUrl: "https://thisisignored.example.com/",
          height: 510,
          width: 450,
          title: "URL Dialog",
        }
      }
    });
  }

  private createCardTaskModuleMEResponse(): Promise<MessagingExtensionActionResponse> {
    return Promise.resolve({
      task: {
        type: 'continue',
        value: {
          card: adaptiveCardBotJson,
          height: 510,
          width: 450,
          title: "Adaptive Card Dialog",
        }
      }
    });
  }

  // Called when the user selects a commmand from the ME's command list when activating an ME
  protected override handleTeamsMessagingExtensionFetchTask(_context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    console.log(`ME FETCH TASK. Action: ${JSON.stringify(action)}`);

    switch (action.commandId) {
      case "fetchCardDialog" : {
        return this.createCardTaskModuleMEResponse();
      }
      case "fetchUrlDialog" : {
        return this.createUrlTaskModuleMEResponse();
      }
      default : {
        console.log(`Unknown commandId: ${action.commandId}`);
      }
    }

    return;
  }

  public override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`TASK MODULE FETCH. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return;
  }

  public override handleTeamsTaskModuleSubmit(_context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`HANDLING DIALOG SUBMIT. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return;
  }
}
