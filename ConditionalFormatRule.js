module.exports = { 
    requests: [
    {
      addConditionalFormatRule: {
        rule: {
          ranges: [
            {
              sheetId: 0,
              startRowIndex: 0,
              startColumnIndex: 6,
              endColumnIndex: 7
            }
          ],
          booleanRule: {
            condition: {
              type: 'TEXT_CONTAINS',
              values: [
                {
                  userEnteredValue: 'Reprovado'
                }
              ]
            },
            format: {
              backgroundColor: {
                red: 0.95,
                green: 0.78,
                blue: 0.76,
              }
            }
          }
        },
        index: 0
      }
    },
    {
      addConditionalFormatRule: {
        rule: {
          ranges: [
            {
              sheetId: 0,
              startRowIndex: 0,
              startColumnIndex: 6,
              endColumnIndex: 7
            }
          ],
          booleanRule: {
            condition: {
              type: 'TEXT_CONTAINS',
              values: [
                {
                  userEnteredValue: 'Aprovado'
                }
              ]
            },
            format: {
              backgroundColor: {
                red: 0.71,
                green: 0.88,
                blue: 0.80,
              }
            }
          }
        },
        index: 0
      }
    },
    {
      addConditionalFormatRule: {
        rule: {
          ranges: [
            {
              sheetId: 0,
              startRowIndex: 0,
              startColumnIndex: 6,
              endColumnIndex: 7
            }
          ],
          booleanRule: {
            condition: {
              type: 'TEXT_CONTAINS',
              values: [
                {
                  userEnteredValue: 'Exame Final'
                }
              ]
            },
            format: {
              backgroundColor: {
                red: 1.0,
                green: 0.91,
                blue: 0.7,
              }
            }
          }
        },
        index: 0
      }
    },
    {
      repeatCell: {
        range: {
          sheetId: 0,
          startRowIndex: 2,
          startColumnIndex: 2,
          endColumnIndex: 8
        },
        cell: {
          userEnteredFormat: {
            horizontalAlignment : 'CENTER',
          }
        },
        fields: 'userEnteredFormat(horizontalAlignment)'
      }
    }
  ]
}