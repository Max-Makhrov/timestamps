# timestamps

Automatic timestamps for Google Sheets.

Try it now, make a copy of the file with the bound script:

https://docs.google.com/spreadsheets/d/1rV60zGQqIoDedp2wpGrTyEY4xcuXil_gTeIcyguMjAE/copy

![Timestams are inserted automatically when a user changes dependent columns](https://raw.githubusercontent.com/Max-Makhrov/timestamps/master/pics/timestamps_teaser%2005.gif)

## Installstion

 1. Change settings on sheet `_onEditSets_`
 2. Go to menu `Timistamp > Get the code!`
 3. Copy the code from the pop-up
 4. Go to menu `Tools > Script Editor...`
 5. Paste the code in the settings section

Insert the settigs between 

```
      ///////////// Change \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
      
      sheetName: ["Sales", "Sales", "Sales"],
      firstRow: ["2", "2", "2"],
      timeStampColumn: ["A", "B", "C"],
      changeColumns: ["D", "H", "D:G;I:"],
      triggerType: ["unique number", "create", "update"],
      insertType: ["value", "value", "value"],
      lastRow: ["-1", "-1", "-1"]
      
      
      ///////////// Change /////////////////////////////////////  
```

<!--stackedit_data:
eyJoaXN0b3J5IjpbLTE3ODAzNDYxNjNdfQ==
-->