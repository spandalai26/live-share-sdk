/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import {
    LivePresence,
    LivePresenceUser,
    LiveState,
    PresenceState,
    UserMeetingRole,
} from "@microsoft/live-share";
import { getRandomDiceValue, stylizeDiceElem } from "./utils";
import { IFluidContainer } from "fluid-framework";
import { AppTheme, IPresenceData } from "./types-interfaces";

export async function renderMeetingStage(
    container: IFluidContainer,
    elem: HTMLElement,
    theme: AppTheme
) {
    const stageTemplate = document.createElement("template");
    stageTemplate["innerHTML"] = `
    <div class="wrapper ${theme} stage">
        <div class="dice"></div>
        <button class="roll">Roll</button>
        <div class="divider"></div>
        <h2>Users:</h2>
    </div>
    `;
    const { diceState, presence } = container.initialObjects;

    elem.appendChild(stageTemplate.content.cloneNode(true));
    const wrapperElem = elem.querySelector<HTMLElement>(".wrapper")!;
    await renderSharedDice(diceState as LiveState<number>, wrapperElem);
    await renderPresenceDiceList(
        presence as LivePresence<IPresenceData>,
        wrapperElem
    );
}

async function renderSharedDice(
    diceState: LiveState<number>,
    wrapperElem: HTMLElement
) {
    const rollButton = wrapperElem.querySelector<HTMLButtonElement>(".roll")!;
    const diceElem = wrapperElem.querySelector<HTMLElement>(".dice")!;

    // Set the value at our dataKey with a random number between 1 and 6.
    rollButton.onclick = () => diceState.set(getRandomDiceValue());

    // Get the current value of the shared data to update the view whenever it changes.
    const updateDice = () => {
        const diceValue = diceState.state;
        stylizeDiceElem(diceElem, diceValue);
    };

    // Use the changed event to trigger the rerender whenever the value changes.
    diceState.on("stateChanged", updateDice);

    // Initialize diceState with initial state of 1 and allowed roles (optional)
    const allowedRoles = [UserMeetingRole.organizer, UserMeetingRole.presenter];
    await diceState.initialize(1, allowedRoles);

    // Render initial dice value
    updateDice();
}

async function renderPresenceDiceList(
    presence: LivePresence<IPresenceData>,
    wrapperElem: HTMLElement
) {
    // Use the changed event to trigger the rerender whenever the remote value changes.
    presence.on("presenceChanged", (userPresence, local) => {
        renderUserDice(presence, userPresence, local, wrapperElem);
    });

    // Initialize presence with a custom data object
    await presence.initialize({
        diceValue: getRandomDiceValue(),
    });
}

// Render a dice owned by a specific user
async function renderUserDice(
    presence: LivePresence<IPresenceData>,
    userPresence: LivePresenceUser<IPresenceData>,
    local: boolean,
    wrapperElem: HTMLElement
) {
    const userDiceElementId = `user-dice-${userPresence.userId}`;
    let userDiceElement = document.getElementById(userDiceElementId);
    // If the user's state is not offline, remove it from the DOM
    if (userPresence.state !== PresenceState.online) {
        userDiceElement?.remove();
        return;
    }
    // Insert user dice element if it is the first time we have received presence change for this user
    if (!userDiceElement) {
        // Insert user's dice roller into wrapper element
        const diceTemplate = document.createElement("template");
        diceTemplate["innerHTML"] = `
        <div id="${userDiceElementId}" class="user-dice-wrapper">
            <div class="user-name"></div>
            <div class="flex">
                ${local ? '<button class="roll user">Roll</button>' : ""}
                <div class="dice user"></div>
            </div>
        </div>
        `;
        wrapperElem.appendChild(diceTemplate.content.cloneNode(true));
        userDiceElement = document.getElementById(userDiceElementId)!;
    }

    // Insert the user's name into the DOM
    const nameElem = userDiceElement.querySelector<HTMLElement>(".user-name")!;
    const displayName = userPresence.displayName ?? "Unknown";
    nameElem.textContent = local ? "You" : displayName;

    // If local user, set the onclick listener for the roll button
    if (local) {
        const rollButton =
            userDiceElement.querySelector<HTMLButtonElement>(".roll")!;
        rollButton.onclick = () =>
            presence.update({
                diceValue: getRandomDiceValue(),
            });
    }

    // Update the dice value in the DOM
    const diceElem = userDiceElement.querySelector<HTMLElement>(".dice")!;
    stylizeDiceElem(diceElem, userPresence.data?.diceValue ?? 1);
}
