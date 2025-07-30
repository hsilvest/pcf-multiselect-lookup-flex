/*
 * MultiSelect Lookup Flex – PCF Control
 *
 * This file implements a simplified multi‑select lookup control for
 * model‑driven Power Apps.  The control provides an editable UI that
 * allows users to search for and select multiple related records and
 * a read‑only representation that displays the selected names as a
 * comma‑separated string.  In this sample the lookup search is
 * simulated using an in‑memory list of items; you can replace it
 * with calls to the Dataverse Web API or the Power Apps component
 * framework lookup API.
 */

import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class MultiSelectLookupFlex implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _context: ComponentFramework.Context<IInputs>;
    private _container: HTMLDivElement;
    private _input: HTMLInputElement;
    private _chipContainer: HTMLDivElement;
    private _selected: { id: string; name: string }[];
    private _notifyOutputChanged: () => void;
    private _isViewMode: boolean;

    /**
     * Called when the control is initialized.  We create DOM elements here
     * and set up event handlers.  The framework will call updateView
     * immediately after init.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;
        this._container = container;
        this._selected = [];

        // Determine if we are in view mode by checking whether the
        // control is disabled.  In a view the control will be read‑only.
        this._isViewMode = !!context.mode.isReadOnly;

        // Create UI elements only if we’re not in view mode.  In view mode
        // we’ll render text in updateView instead.
        if (!this._isViewMode) {
            this._chipContainer = document.createElement("div");
            this._chipContainer.className = "ms-multiselect-chip-container";
            this._container.appendChild(this._chipContainer);

            this._input = document.createElement("input");
            this._input.type = "text";
            this._input.className = "ms-multiselect-input";
            this._input.placeholder = "Search...";
            this._input.addEventListener("keydown", this.onInputKeyDown.bind(this));
            this._container.appendChild(this._input);
        }
    }

    /**
     * Event handler for key down events on the search input.  When the
     * user presses Enter, we simulate a lookup by adding a new item
     * from a predefined list that matches the search term.  If no
     * match is found, nothing happens.
     */
    private onInputKeyDown(evt: KeyboardEvent): void {
        if (evt.key !== "Enter" || !this._input.value) {
            return;
        }

        const term = this._input.value.toLowerCase();
        // Simulated dataset.  Replace this with an API call.
        const dataset: { id: string; name: string }[] = [
            { id: "1", name: "Contoso" },
            { id: "2", name: "Fabrikam" },
            { id: "3", name: "Adventure Works" },
            { id: "4", name: "Northwind" }
        ];
        const match = dataset.find(item => item.name.toLowerCase().startsWith(term));
        if (match && !this._selected.some(s => s.id === match.id)) {
            this._selected.push(match);
            this.addChip(match);
            this._notifyOutputChanged();
        }
        // Clear input after selection.
        this._input.value = "";
    }

    /**
     * Creates a visual chip representing a selected item.  Each chip has
     * a remove button that allows the user to deselect the item.
     */
    private addChip(item: { id: string; name: string }): void {
        const chip = document.createElement("span");
        chip.className = "ms-multiselect-chip";
        chip.textContent = item.name;

        const removeBtn = document.createElement("button");
        removeBtn.textContent = "×";
        removeBtn.className = "ms-multiselect-remove";
        removeBtn.addEventListener("click", () => {
            this._selected = this._selected.filter(s => s.id !== item.id);
            this._chipContainer.removeChild(chip);
            this._notifyOutputChanged();
        });
        chip.appendChild(removeBtn);
        this._chipContainer.appendChild(chip);
    }

    /**
     * Called when the framework wants to notify the control that
     * properties have changed.  We use this as an opportunity to
     * update the read‑only rendering.
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const rawValue: string | undefined = context.parameters.value.raw || "";
        if (this._isViewMode) {
            // In view mode we simply set the inner text to the stored value.
            this._container.innerText = rawValue
                .split(";")
                .filter(v => v)
                .map(v => v.trim())
                .join(", ");
        }
    }

    /**
     * Returns the control outputs to the framework.  We join selected
     * item IDs with a semicolon; you might choose to store a JSON
     * representation instead.
     */
    public getOutputs(): IOutputs {
        const ids = this._selected.map(s => s.id);
        return { value: ids.join(";") };
    }

    /**
     * Called when the control is removed from the DOM.  Use this to
     * clean up event handlers and DOM nodes that were created in
     * init().
     */
    public destroy(): void {
        // Clean up event handlers to avoid memory leaks.
        if (this._input) {
            this._input.removeEventListener("keydown", this.onInputKeyDown);
        }
        while (this._container.firstChild) {
            this._container.removeChild(this._container.firstChild);
        }
    }
}
