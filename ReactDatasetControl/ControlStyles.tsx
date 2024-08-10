import { IStackTokens, IStackStyles, IDetailsListStyles, IDropdownStyles } from "@fluentui/react";

export const stackTokens: IStackTokens = { 
    childrenGap: 10 
};

export const stackStyles: IStackStyles = {
    root: {
        padding: 10,
        width: '100%',
        marginBottom: 20,
    },
};

export const detailsListStyles: IDetailsListStyles = {
    root: {
        overflowX: 'auto'
    }, 
    contentWrapper: {
        overflowY: 'auto', 
        width: 'max-content', 
        height: 450
    },
    focusZone : {},
    headerWrapper: {} 
}

export const dropDownStyles : Partial<IDropdownStyles> = {
    root : {
        width: 'auto', 
        minWidth: 200
    }
}