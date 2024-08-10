import * as React from 'react';
import { useState, useEffect } from 'react';
import { IInputs } from "./generated/ManifestTypes";
import { DetailsList, DetailsListLayoutMode, Dropdown, IDropdownOption, PrimaryButton, Stack } from '@fluentui/react';
import * as XLSX from 'xlsx';
import {stackTokens, stackStyles, detailsListStyles, dropDownStyles} from './ControlStyles'
import { text } from 'stream/consumers';

export function PCFControl({sampleDataSet} : IInputs) {

    function parseDatasetToJSON() 
    {
        const jsonData = [];
        for(const recordID of sampleDataSet.sortedRecordIds) 
        {
            // Dataset record
            const record = sampleDataSet.records[recordID];
            const temp: Record<string, any> = {};

            for(const column of sampleDataSet.columns) 
            {
                temp[column.name] = record.getFormattedValue(column.name)
            }
            jsonData.push(temp);
        }
        return jsonData;
    }

    function createFileOptions(notes: Array<Record<string, any>>)
    {
        const options: IDropdownOption[] = [];

        for(const [index, note] of notes.entries())
        {
            const option = { key: index, text: note["filename"] ?? "No File" };
            options.push(option);
        }
        return options;
    }

    function createSheetOptions(workbook: XLSX.WorkBook)
    {
        const options: IDropdownOption[] = [];
        for(const [index, sheetName] of workbook.SheetNames.entries())
        {
            const option = { key: index, text: sheetName };
            options.push(option);
        }
        return options;
    }

    function handleSelectFile(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption)
    {
        if(option === undefined) return; // Return if no option is selected

        const note = notes[option.key as number]; // Get note record using index
        const base64Data = note["documentbody"] ?? ""; // Get file data
        const workbook = XLSX.read(base64Data, { type: 'base64', cellDates: true }); // Converts base64 data to excel workbook object
        setExcelWorkbook(workbook);

        const sheetOptionsRecords = createSheetOptions(workbook);
        setSheetOptions(sheetOptionsRecords);
    }

    function handleSelectSheet(event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption)
    {
        if(option === undefined) return; // Return if no option is selected

        const sheet = excelWorkbook.Sheets[option.text as string]; // Get sheet record using SheetName
        const rowRecords: Record<string, any>[] = XLSX.utils.sheet_to_json(sheet, {raw: false}); // Sheet Records in JSON Array
        setRows(rowRecords);
    }

    function handleDownload()
    {
        XLSX.writeFile(excelWorkbook, 'download.xlsx');
    }

    const [notes, setNotes] = useState<Array<Record<string, any>>>([]); // Array of JSON objects
    const [fileOptions, setFileOptions] = useState<IDropdownOption[]>([]); // Array of IDropdownOption objects
    const [excelWorkbook, setExcelWorkbook] = useState<XLSX.WorkBook>(XLSX.utils.book_new()); // Excel workbook object
    const [sheetOptions, setSheetOptions] = useState<IDropdownOption[]>([]); // Array of IDropdownOption objects
    const [rows, setRows] = useState<Array<Record<string, any>>>([]); // Array of JSON objects
    

    useEffect(() => {
        const notesRecords = parseDatasetToJSON();
        const fileOptionsRecords = createFileOptions(notesRecords);

        setNotes(notesRecords);
        setFileOptions(fileOptionsRecords);

    }, [sampleDataSet]); // On dataset change

    

    return (
        <Stack tokens={stackTokens} styles={stackStyles}>
            <Stack horizontal tokens={stackTokens} styles={stackStyles}>
                <Dropdown 
                    placeholder="Select File" 
                    options={fileOptions} 
                    onChange={handleSelectFile} 
                    styles={dropDownStyles}
                />
                <Dropdown 
                    placeholder="Select Sheet" 
                    options={sheetOptions} 
                    onChange={handleSelectSheet} 
                    styles={dropDownStyles}
                    defaultSelectedKey='0' 
                />
                <PrimaryButton text="Download" allowDisabledFocus onClick={handleDownload}/>
            </Stack>
            <DetailsList 
                items={rows} 
                styles={detailsListStyles}
                data-is-scrollable={false}
                layoutMode={DetailsListLayoutMode.fixedColumns}
            />
        </Stack>
    );
}