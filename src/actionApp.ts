import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionAction,
  MessagingExtensionActionResponse,
  TaskModuleRequest,
  TaskModuleResponse,
  ActionTypes,
} from "botbuilder";

const urlDialogTriggerValue = "requestUrl";
const cardDialogTriggerValue = "requestCard";
const messagePageTriggerValue = "requestMessage";
const noResponseTriggerValue = "requestNoResponse";

const pageDomain = "localhost:53000";
// const pageDomain = "helloworld36cffe.z5.web.core.windows.net";

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
    } else if (commandData.data === messagePageTriggerValue) {
      return this.createMessagePageMEResponse();
    }

    if (commandId === "createCard" && commandData.title === "config") {
      return {
        composeExtension: {
          type: "config",
          suggestedActions: {
            actions: [
                {
                  title: "Config Action Title",
                  type: ActionTypes.OpenUrl,
                  value: `https://${pageDomain}/index.html?page=config#/tab`
                },
            ],
        },
      },
      };
    }
   
    if (commandId === "createCard" && commandData.title === "auth") {
      return {
        composeExtension: {
          type: "auth",
          suggestedActions: {
            actions: [
                {
                  title: "Auth Action Title",
                  type: ActionTypes.OpenUrl,
                  value: `https://${pageDomain}/index.html?page=auth#/tab`
                },
            ],
          },
        },
      };
    }

    if (commandId === "createCard" && commandData.title === "sso") {
      return {
        composeExtension: {
          type: 'silentAuth',
          suggestedActions: {
              actions: [
                  {
                    title: "SSO?",
                    type: ActionTypes.OpenUrl,
                    value: `https://${pageDomain}/index.html?page=auth#/tab`
                  },
              ],
          },
        },
      };
    }

    switch (commandId) {
      case "createCard":
        {
          const attachment = CardFactory.heroCard(
            `${commandData.title} - ${commandData.subTitle}`,
            commandData.text,
            null, // no images
            [{
              type: 'invoke',
              title: "Show URL Task Module",
              value: {
                  type: 'task/fetch',
                  data: urlDialogTriggerValue
              }
            },
            {
              type: 'invoke',
              title: "Show Adaptive Card Task Module",
              value: {
                  type: 'task/fetch',
                  data: cardDialogTriggerValue
              }
            }]
          );
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
          url: `https://${pageDomain}/index.html?randomNumber=${this.getRandomIntegerBetween(1, 1000)}#/tab`,
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

  private createMessagePageMEResponse(): Promise<MessagingExtensionActionResponse> {
    return Promise.resolve({
      task: {
          type: 'message',
          value: `Hello! This is a message!`,
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
      case "triggerConfigPage" : {
        return Promise.resolve({
          composeExtension: {
            type: "config",
            suggestedActions: {
              actions: [
                  {
                    title: "Trigger config page",
                    type: ActionTypes.OpenUrl,
                    value: `https://${pageDomain}/index.html?page=config#/tab`
                  },
              ],
            },
          },
        });
      }
      case "triggerOAuthPage" : {
        return Promise.resolve({
          composeExtension: {
            type: "auth",
            suggestedActions: {
              actions: [
                  {
                    title: "Auth Action Title",
                    type: ActionTypes.OpenUrl,
                    value: `https://${pageDomain}/index.html?page=auth#/tab`
                  },
              ],
            },
          },
        });
      }
      case "triggerSsoPage" : {
        return Promise.resolve({
          composeExtension: {
            type: 'silentAuth',
            suggestedActions: {
                actions: [
                    {
                      title: "SSO?",
                      type: ActionTypes.OpenUrl,
                      value: `https://${pageDomain}/index.html?page=auth#/tab`
                    },
                ],
            },
          },
        });
      }
      default : {
        console.log(`Unknown commandId: ${action.commandId}`);
      }
    }

    return;
  }

  private createResponseToTaskModuleRequest(taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    const taskRequestData = taskModuleRequest.data.data;

    switch (taskRequestData) {
      case urlDialogTriggerValue:
        return this.createUrlTaskModuleMEResponse();

      case cardDialogTriggerValue:
        return this.createCardTaskModuleMEResponse();

      case messagePageTriggerValue:
        return this.createMessagePageMEResponse();

      case noResponseTriggerValue:
        return;

      default:
        return Promise.resolve({
          task: {
              type: 'message',
              value: `The submitted data did not contain a valid request (submitted data: ${taskModuleRequest.data})`,
          }
      });
    }
  }  
  
  public override handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`TASK MODULE FETCH. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return this.createResponseToTaskModuleRequest(taskModuleRequest);
  }

  public override handleTeamsTaskModuleSubmit(_context: TurnContext, taskModuleRequest: TaskModuleRequest): Promise<TaskModuleResponse> {
    console.log(`HANDLING DIALOG SUBMIT. Task module request: ${JSON.stringify(taskModuleRequest)}`);

    return this.createResponseToTaskModuleRequest(taskModuleRequest);
  }

  public override handleTeamsMessagingExtensionCardButtonClicked(_context: TurnContext, cardData: any): Promise<void> {
    console.log(`HANDLING CARD BUTTON CLICKED. Card data: ${JSON.stringify(cardData)}`);

    return;
  }

}
