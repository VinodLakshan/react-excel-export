import logo from './logo.svg';
import './App.css';
import * as XLSX from 'xlsx-js-style';

function App() {

  const users = [
    {
      "userId": 1,
      "firstName": "Krish",
      "lastName": "Lee",
      "phoneNumber": "123456",
      "emailAddress": "krish.lee@learningcontainer.com"
    },
    {
      "userId": 2,
      "firstName": "racks",
      "lastName": "jacson",
      "phoneNumber": "123456",
      "emailAddress": "racks.jacson@learningcontainer.com"
    },
    {
      "userId": 3,
      "firstName": "denial",
      "lastName": "roast",
      "phoneNumber": "33333333",
      "emailAddress": "denial.roast@learningcontainer.com"
    },
    {
      "userId": 4,
      "firstName": "devid",
      "lastName": "neo",
      "phoneNumber": "222222222",
      "emailAddress": "devid.neo@learningcontainer.com"
    },
    {
      "userId": 5,
      "firstName": "jone",
      "lastName": "mac",
      "phoneNumber": "111111111",
      "emailAddress": "jone.mac@learningcontainer.com"
    }
  ];

  const headers = ["User Id", "First Name", "Last Name", "Phone Number", "Email Address"];

  const writeToExcel = () => {

    const sheetData = getSheetData(headers, users);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    setSheetStyles(ws, sheetData);

    XLSX.utils.book_append_sheet(wb, ws, "Users");
    // XLSX.writeFile(wb, "all_users.xlsx")

  }

  const getSheetData = (headers, data) => {

    const fields = Object.keys(data[0]);
    const sheetData = data.map(row => {
      return fields.map(fieldName => {
        return row[fieldName] ? row[fieldName] : "";
      });
    });

    sheetData.unshift(headers);
    console.log(sheetData)
    return sheetData;
  }

  const setSheetStyles = (ws, data) => {

    const cols = [];

    // headers
    for (let i in data[0]) {

      const col = XLSX.utils.encode_col(i);

      ws[col + 1].s = {
        font: {
          bold: true,
          color: { rgb: "FFFFFF" },
        },
        fill: {
          patternType: "solid",
          fgColor: { rgb: "00AB55" }
        },
        alignment: {
          vertical: "center",
          horizontal: "center"
        }
      };

      // col width
      cols.push({ wch: 25 })
    }

    // data
    for (let i = 1; i < data.length; i++) {

      const row = XLSX.utils.encode_row(i);

      for (let j in data[i]) {

        const col = XLSX.utils.encode_col(j);

        ws[col + row].s = {
          alignment: {
            vertical: "center",
            horizontal: "center"
          }
        };
      }

    }

    ws['!cols'] = cols;
  }

  return (
    <div className="App">
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <button
          onClick={writeToExcel}
          style={{ backgroundColor: 'white', padding: "10px", cursor: "pointer" }}>
          Download data as an excel file
        </button>
      </header>
    </div>
  );
}

export default App;
