import * as React from "react";
import styles from "./AdaptiveCardHost.module.scss";
import {
  IAdaptiveCardHostProps,
  IAdaptiveCardHostState,
} from "./IAdaptiveCardHostProps";

// Needed for the placeholder when the web part is not configured
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

// Needed for cycling through array dataSource
import {
  Carousel,
  CarouselButtonsLocation,
  CarouselButtonsDisplay,
} from "@pnp/spfx-controls-react/lib/Carousel";
// import { sp } from "@pnp/sp";

// Needed for displaying adaptive card results
import {
  AdaptiveCard,
  IAdaptiveCardActionResult,
} from "../../../controls/AdaptiveCard";

// Needed for displaying warnings
import {
  MessageBar,
  MessageBarType,
  MessageBarButton,
  arraysEqual,
} from "office-ui-fabric-react";

// Localization
import * as strings from "AdaptiveCardHostWebPartStrings";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

export default class AdaptiveCardHost extends React.Component<
  IAdaptiveCardHostProps,
  IAdaptiveCardHostState,
  {}
> {
  constructor(props: IAdaptiveCardHostProps, state: IAdaptiveCardHostState) {
    super(props);
    this.state = {
      dataLength: this.props.data ? JSON.parse(this.props.data).length - 1 : 0,
      currentIndex: 0,
      // WindowSize:this.props.WebpartElement.getBoundingClientRect().width
    };
  }

  /**
   * Renders the adaptive card, or one of the many warnings
   */
  public render(): React.ReactElement<IAdaptiveCardHostProps> {
    const {
      template,
      data,
      useTemplating,
      useArrayCycling,
      themeVariant,
    } = this.props;

    // if we didn't specify a template, we need a template!
    const needsTemplate: boolean = !template;

    // If we use Adaptive Card Templating and didn't specify data, we need data!
    const needsData: boolean = useTemplating && !data;

    // If we didn't use Adaptive Card Templating but the template contains $data nodes,
    // if means it is a data-enabled template
    const dataEnabledTemplate: boolean =
      template && template.indexOf('"$data"') > -1;

    // If we didn't specify the template, show the placeholder
    if (needsTemplate) {
      return (
        <Placeholder
          iconName="Code"
          iconText={strings.PlaceholderIconText}
          description={strings.PlaceholderDescription}
          buttonLabel="Configure"
          onConfigure={this._configureHandler}
        />
      );
    } else if (needsData) {
      // If we didn't specify data and we need it, display a different placeholder
      return (
        <Placeholder
          iconName="PageData"
          iconText={strings.DataNeededIconText}
          description={strings.DataNeededDescription}
          buttonLabel={strings.DataNeededButtonLabel}
          onConfigure={this._configureHandler}
        />
      );
    } else {
      // this.setState({
      //   dataLength: this.props.data ? this.props.data.length : 0,
      // });

      // Display the Adaptive Card
      return (
        <>
          {dataEnabledTemplate && !useTemplating && (
            <MessageBar
              dismissButtonAriaLabel="Close"
              messageBarType={MessageBarType.warning}
              actions={
                <div>
                  <MessageBarButton onClick={this._configureHandler}>
                    {strings.ConfigureButtonLabel}
                  </MessageBarButton>
                </div>
              }
            >
              {strings.AdaptingTemplatingWarningIntro}
              <a
                href={strings.AdaptiveCardTemplatingMoreInfoLinkUrl}
                target="_blank"
              >
                {strings.AdaptiveCardTemplating}
              </a>
              {strings.AdaptiveCardWarningPartTwo}
              <strong>{strings.UseAdaptiveTemplatingLabel}</strong>
              {strings.AdaptiveTemplatingEnd}
            </MessageBar>
          )}

          {useArrayCycling === true &&
            Array.isArray(JSON.parse(this.props.data)) && (
              <Carousel
                buttonsLocation={CarouselButtonsLocation.top}
                buttonsDisplay={CarouselButtonsDisplay.block}
                isInfinite={true}
                element={
                  <AdaptiveCard
                    template={template}
                    data={JSON.stringify(
                      JSON.parse(this.props.data)[this.state.currentIndex]
                    )}
                    useTemplating={useTemplating}
                    useArrayCycling={useArrayCycling}
                    themeVariant={themeVariant}
                    onExecuteAction={this._executeActionHandler}
                    className={styles.adaptiveCardHost}
                  />
                }
                onMoveNextClicked={(index: number) => {
                  let _currentIndex = this.state.currentIndex + 1;
                  if (_currentIndex > this.state.dataLength) _currentIndex = 0;

                  this.setState({ currentIndex: _currentIndex });
                  console.log(`Next : ${_currentIndex} -> ${index}`);
                }}
                onMovePrevClicked={(index: number) => {
                  let _currentIndex = this.state.currentIndex - 1;
                  if (isNaN(_currentIndex) || _currentIndex < 0)
                    _currentIndex = this.state.dataLength;

                  this.setState({ currentIndex: _currentIndex });
                  console.log(`Prev : ${_currentIndex} -> ${index}`);
                }}
              />
            )}

          {!useArrayCycling && (
            <AdaptiveCard
              template={template}
              data={data}
              useTemplating={useTemplating}
              useArrayCycling={useArrayCycling}
              themeVariant={themeVariant}
              onExecuteAction={this._executeActionHandler}
              className={styles.adaptiveCardHost}
            />
          )}
        </>
      );
    }
  }

  /**
   * Demonstrates how we can respond to actions
   */
  private _executeActionHandler = (action: IAdaptiveCardActionResult) => {
    console.log("Action", action);

    // Feel free to handle actions any way you want
    switch (action.type) {
      case "Action.OpenUrl":
        window.open(action.url);
        break;
      case "Action.Submit":
        alert(action.title);
        break;
      default:
        break;
    }
  };

  /** Opens the configuration pane */
  private _configureHandler = () => {
    this.props.context.propertyPane.open();
  };
}
