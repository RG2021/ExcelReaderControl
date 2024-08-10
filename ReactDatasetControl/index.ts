import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { HelloWorld, IHelloWorldProps } from "./HelloWorld";
import { PCFControl } from "./App"
import * as React from "react";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

export class ReactDatasetControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;
    constructor() {}
    
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary): void {
        this.notifyOutputChanged = notifyOutputChanged;
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {

        if (context.parameters.sampleDataSet.addColumn) {
            context.parameters.sampleDataSet.addColumn("documentbody");
            context.parameters.sampleDataSet.addColumn("filename");
            context.parameters.sampleDataSet.refresh();
        }
        
        const props: IInputs = { sampleDataSet: context.parameters.sampleDataSet };
        return React.createElement(PCFControl, props);
    }

    public getOutputs(): IOutputs {
        return { };
    }

    public destroy(): void {}
}
