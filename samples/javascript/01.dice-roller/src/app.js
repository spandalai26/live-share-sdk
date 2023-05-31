/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { SharedMap } from "fluid-framework";
import { app, pages, LiveShareHost } from "@microsoft/teams-js";
import { LiveShareClient, TestLiveShareHost } from "@microsoft/live-share";
import { InsecureTokenProvider } from "@fluidframework/test-client-utils";

<<<<<<< HEAD
const searchParams = new URL(window.location).searchParams;
=======
const searchParams = new URL(window.location.href).searchParams;
const IN_TEAMS = searchParams.get("inTeams") === "1";
>>>>>>> bcdddf8bf71aea1b22d95ef52ecbd2451acb64b6
const root = document.getElementById("content");

// Define container schema

const diceValueKey = "dice-value-key";

const containerSchema = {
  initialObjects: { diceMap: SharedMap },
};

function onContainerFirstCreated(container) {
  // Set initial state of the rolled dice to 1.
  container.initialObjects.diceMap.set(diceValueKey, 1);
}

// STARTUP LOGIC

async function start() {
  // Check for page to display
  let view = searchParams.get("view") || "stage";

<<<<<<< HEAD
  // Check if we are running on stage.
  if (!!searchParams.get("inTeams")) {
    // Initialize teams app
    await app.initialize();
=======
    // Check if we are running on stage.
    if (IN_TEAMS) {
        // Initialize teams app
        await app.initialize();
        // Get Teams app context to get the initial theme
        const context = await app.getContext();
        theme = context.app.theme === "default" ? "light" : "dark";
        app.registerOnThemeChangeHandler((theme) => {
            theme = theme === "default" ? "light" : "dark";
        });
    }
>>>>>>> bcdddf8bf71aea1b22d95ef52ecbd2451acb64b6

    // Get our frameContext from context of our app in Teams
    const context = await app.getContext();
    if (context.page.frameContext == "meetingStage") {
      view = "stage";
    }
  }

  // Load the requested view
  switch (view) {
    case "content":
      renderSidePanel(root);
      break;
    case "config":
      renderSettings(root);
      break;
    case "stage":
    default:
      const { container } = await joinContainer();
      renderStage(container.initialObjects.diceMap, root);
      break;
  }
}

start().catch((error) => console.error(error));

async function joinContainer() {
<<<<<<< HEAD
    // Are we running in teams? If so, use LiveShareHost, otherwise use TestLiveShareHost
    const host = !!searchParams.get("inTeams")
      ? LiveShareHost.create()
      : TestLiveShareHost.create();
=======
    // Are we running in teams?
    const host = IN_TEAMS ? LiveShareHost.create() : TestLiveShareHost.create();

>>>>>>> bcdddf8bf71aea1b22d95ef52ecbd2451acb64b6
    // Create client
    const liveShare = new LiveShareClient(host);
    // Join container
    return await liveShare.joinContainer(containerSchema, onContainerFirstCreated);
  }


const stageTemplate = document.createElement("template");

stageTemplate["innerHTML"] = `
  <div class="wrapper">
    <div class="dice"></div>
    <button class="roll"> Roll </button>
  </div>
`;
function renderStage(diceMap, elem) {
  elem.appendChild(stageTemplate.content.cloneNode(true));
  const rollButton = elem.querySelector(".roll");
  const dice = elem.querySelector(".dice");

  rollButton.onclick = () => updateDice(Math.floor(Math.random() * 6) + 1);

  const updateDice = (value) => {
    // Unicode 0x2680-0x2685 are the sides of a die (⚀⚁⚂⚃⚄⚅).
    dice.textContent = String.fromCodePoint(0x267f + value);
  };
  updateDice(1);
}

<<<<<<< HEAD

rollButton.onclick = () =>
  diceMap.set("dice-value-key", Math.floor(Math.random() * 6) + 1);


const updateDice = () => {
  const diceValue = diceMap.get("dice-value-key");
  dice.textContent = String.fromCodePoint(0x267f + diceValue);
};
updateDice();

diceMap.on("valueChanged", updateDice);


const sidePanelTemplate = document.createElement("template");


sidePanelTemplate["innerHTML"] = `
  <style>
    .wrapper { text-align: center }
    .title { font-size: large; font-weight: bolder; }
    .text { font-size: medium; }
  </style>
  <div class="wrapper">
    <p class="title">Lets get started</p>
    <p class="text">Press the share to stage button to share Dice Roller to the meeting stage.</p>
  </div>
`;

function renderSidePanel(elem) {
  elem.appendChild(sidePanelTemplate.content.cloneNode(true));
}

const settingsTemplate = document.createElement("template");

settingsTemplate["innerHTML"] = `
  <style>
    .wrapper { text-align: center }
    .title { font-size: large; font-weight: bolder; }
    .text { font-size: medium; }
  </style>
  <div class="wrapper">
    <p class="title">Welcome to Dice Roller!</p>
    <p class="text">Press the save button to continue.</p>
  </div>
`;

function renderSettings(elem) {
  elem.appendChild(settingsTemplate.content.cloneNode(true));

  // Save the configurable tab
  pages.config.registerOnSaveHandler((saveEvent) => {
    pages.config.setConfig({
      websiteUrl: window.location.origin,
      contentUrl: window.location.origin + "?inTeams=1&view=content",
      entityId: "DiceRollerFluidLiveShare",
      suggestedDisplayName: "DiceRollerFluidLiveShare",
    });
    saveEvent.notifySuccess();
  });

  // Enable the Save button in config dialog
  pages.config.setValidityState(true);
}
=======
start().catch((error) => renderError(root, error, theme));
>>>>>>> bcdddf8bf71aea1b22d95ef52ecbd2451acb64b6
