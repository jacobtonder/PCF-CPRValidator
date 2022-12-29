import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class PCFCPRValidator implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;
    private _customAttributeElement: HTMLInputElement;
    private _errorElement: HTMLElement;
    private _customAttributeChanged: EventListenerOrEventListenerObject;
    private _customAttributeValue: string;

    /**
     * Empty constructor.
     */
    constructor()
    {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement): void
    {
        this._context = context; 
        this._notifyOutputChanged = notifyOutputChanged; 
        this._container = container;
        this._customAttributeChanged = this.customAttributeChanged.bind(this);

        // Design custom attribute Html Input Element
        this._customAttributeElement = document.createElement("input");
        this._customAttributeElement.setAttribute("type", "text");
        this._customAttributeElement.setAttribute("class", "classInput");
        this._customAttributeElement.addEventListener("change", this._customAttributeChanged);

        // Check if Control is Disabled/Read-Only
        var disabled = context.mode.isControlDisabled;
        this._customAttributeElement.disabled = disabled; 

        var currentValue = context.parameters.value.raw ? context.parameters.value.raw : "";
        this._customAttributeElement.setAttribute("value", currentValue);
        this._customAttributeElement.value = currentValue;

        // Add an error visual to show the error message when there is an invalid credit card
        this._errorElement = document.createElement("div");
        this._errorElement.setAttribute("class", "controlDiv");
        var errorChild1 = document.createElement("label");
        errorChild1.setAttribute("class", "controlImage")
        errorChild1.innerText = " This is not a valid CPR";
        
        var errorChild2 = document.createElement("label");
        errorChild2.setAttribute("class", "controlLabel")
        errorChild2.innerText = "";
        
        this._errorElement.appendChild(errorChild1);
        this._errorElement.appendChild(errorChild2);

        this._container.appendChild(this._customAttributeElement);
        this._container.appendChild(this._errorElement);

        if (this.isValidCPR(currentValue))
        {
            this._errorElement.style.display = "none";
            this._customAttributeValue = currentValue;
        }
        else
        {
            this._errorElement.style.display = "block";
            this._customAttributeValue = "";
        }
        
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void
    {
        if (context.parameters.value.raw != null)
            this._customAttributeValue = context.parameters.value.raw;

        this._context = context;
        this._customAttributeElement.setAttribute("value", context.parameters.value.formatted 
            ? context.parameters.value.formatted
            : ""
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {
            value: this._customAttributeValue
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        this._customAttributeElement.removeEventListener("change", this._customAttributeChanged);
    }

    private customAttributeChanged(evt: Event): void
    {
        var context = this._context;
        var customAttributeValue = this._customAttributeElement.value;

        if (this.isValidCPR(customAttributeValue))
        {
            this._errorElement.style.display = "none";
            this._customAttributeValue = customAttributeValue;
        }
        else
        {
            this._errorElement.style.display = "block";
            this._customAttributeValue = "";
        }

        this._notifyOutputChanged();
    }

    private isValidCPR(value: string): boolean
    {
        var deliminatorRegex = new RegExp("-", "g");
        var cpr = value.replace(deliminatorRegex, '');
        var cprRegex = /^\d{10}$/;

        if (cpr === null)
            return false;

        if (!cprRegex.test(cpr))
            return false;

        var month = parseInt(cpr.substring(2, 4));
        if(month === 0 || month > 12)
            return false;

        var daysInMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
        var day = parseInt(cpr.substring(0, 2));
        if (day > daysInMonth[month - 1])
            return false;

        return this.isModulus(cpr);
    }

    private isModulus(value: string): boolean
    {
        var components = value.split('');
        var controlNumbers = [4, 3, 2, 7, 6, 5, 4, 3, 2, 1];
        var sum = 0;

        components.forEach((element, index) => {
            var currentControl = controlNumbers[index];
            sum += parseInt(element) * currentControl;
        });

        return sum % 11 === 0;
    }
}
