const SHEET_COLUMNS = {
  A: 1,
  B: 2,
  C: 3,
  D: 4,
  E: 5,
  F: 6,
  G: 7,
  H: 8,
  I: 9,
  J: 10,
  K: 11,
  L: 12,
  M: 13,
  N: 14,
  O: 15,
  P: 16,
  Q: 17,
  R: 18,
  S: 19,
  T: 20,
  U: 21,
  V: 22,
  W: 23,
  X: 24,
  Y: 25,
  Z: 26
};
const SHEET_COLUMNS_ALT = {
  1: "A",
  2: "B",
  3: "C",
  4: "D",
  5: "E",
  6: "F",
  7: "G",
  8: "H",
  9: "I",
  10: "J",
  11: "K",
  12: "L",
  13: "M",
  14: "N",
  15: "O",
  16: "P",
  17: "Q",
  18: "R",
  19: "S",
  20: "T",
  21: "U",
  22: "V",
  23: "W",
  24: "X",
  25: "Y",
  26: "Z"
};

const MAPPING_COLUMNS = {
  Escenario: "testsuite",
  "Descripcion del Escenario": "details",
  "Casos de Prueba": "testcase",
  "Descripcion del Caso de Prueba": "summary",
  "Pre - Condiciones": "preconditions",
  //"Datos de Prueba": "",
  Pasos: "steps"
  //"Resultado Esperado": "",
  //"Post - Condiciones": ""
  //Prioridad: "importance",
  //"Regresion (S/N)": ""
  //"Escenarios Vinculados": "",
  //Comentarios: ""
};
const generateCDATA = value => `<![CDATA[${value}]]>`;
const sanitazeText = value => value;

const convertSheetToObject = (workbook, firstLimit, lastLimit) => {
  const rowLimit = [
    parseInt(firstLimit.replace(/^\D+/g, "")),
    parseInt(lastLimit.replace(/^\D+/g, ""))
  ];
  const columnLimit = [
    SHEET_COLUMNS[firstLimit.replace(rowLimit[0], "")],
    SHEET_COLUMNS[lastLimit.replace(rowLimit[1], "")]
  ];

  const table = workbook.Sheets["Especificacion de CP"];
  const columns = [];
  for (let i = columnLimit[0]; i <= columnLimit[1]; i++) {
    const columnName = table[`${SHEET_COLUMNS_ALT[i]}${rowLimit[0]}`].v;

    const columnValue = MAPPING_COLUMNS[columnName];
    if (columnValue) {
      columns.push({
        columnValue,
        columnName,
        columnLetter: SHEET_COLUMNS_ALT[i]
      });
    }
  }
  const result = [];
  for (let i = rowLimit[0] + 1; i <= rowLimit[1]; i++) {
    const item = {};
    columns.forEach(column => {
      const cellValue = table[`${column.columnLetter}${i}`];
      item[column.columnValue] = cellValue && cellValue.v;
    });
    const testsuiteIndex = result.findIndex(
      a => a.testsuite === item.testsuite
    );

    if (testsuiteIndex >= 0) {
      result[testsuiteIndex].testCases.push(item);
    } else {
      result.push({
        testsuite: item.testsuite,
        details: item.details,
        testCases: [item]
      });
    }
  }
  return result;
};

const convertObjectToXML = escenarios => {
  var XML = new XMLWriter();
  XML.BeginNode("testsuite");
  XML.Attrib("name", "");
  XML.BeginNode("details");
  XML.WriteString(generateCDATA(""));
  XML.EndNode();
  escenarios.forEach(escenario => {
    XML.BeginNode("testsuite");
    XML.Attrib("name", escenario.testsuite);

    XML.BeginNode("details");
    XML.WriteString(generateCDATA(escenario.testsuite));
    XML.EndNode();

    escenario.testCases.forEach(testcase => {
      XML.BeginNode("testcase");
      XML.Attrib("name", testcase.testcase);
      XML.BeginNode("summary");
      XML.WriteString(generateCDATA(testcase.testcase));
      XML.EndNode();
      XML.BeginNode("preconditions");
      XML.WriteString(generateCDATA(testcase.preconditions));
      XML.EndNode("preconditions");
      XML.BeginNode("status");
      XML.WriteString("1");
      XML.EndNode();

      XML.BeginNode("execution_type");
      XML.WriteString(generateCDATA("1"));
      XML.EndNode();

      XML.BeginNode("estimated_exec_duration");
      XML.WriteString(generateCDATA("5"));
      XML.EndNode();

      XML.BeginNode("importance");
      XML.WriteString(generateCDATA("3"));
      XML.EndNode();

      XML.EndNode(); //end of testcase node
    });

    XML.EndNode();
  });
  XML.EndNode();
  XML.Close();
  //console.log("XML", XML.ToString().replace(/</g, "\n<"));
  return XML;
};

const downloadXML = xmltext => {
  const filename = "file.xml";
  const pom = document.createElement("a");
  const bb = new Blob([xmltext], { type: "text/plain" });
  pom.setAttribute("href", window.URL.createObjectURL(bb));
  pom.setAttribute("download", filename);
  pom.dataset.downloadurl = ["text/plain", pom.download, pom.href].join(":");
  pom.draggable = true;
  pom.classList.add("dragout");
  pom.click();
};
