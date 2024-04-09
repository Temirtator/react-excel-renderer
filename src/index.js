import React, { Component } from 'react';
import XLSX from 'xlsx';

export class OutTable extends Component {
  constructor(props) {
    super(props);
    this.rows = props.data;
    this.cols = props.columns;
    this.withZeroColumn = props.withZeroColumn;
    this.withoutRowNum = props.withoutRowNum;
    this.tableHeaderRowClass = props.tableHeaderRowClass;
    this.className = props.className;
    this.tableClassName = props.tableClassName;
    this.renderRowNum = props.renderRowNum;
  }

  renderHeader() {
    return (
      <tr>
        {
          this.withZeroColumn && !this.withoutRowNum && 
            <th className={this.tableHeaderRowClass || ""}></th>
        }
        {
          this.cols.map((c) =>
            <th key={c.key} className={c.key === -1 ? this.tableHeaderRowClass : ""}>{c.key === -1 ? "" : c.name}</th>)
        }
      </tr>
    )
  }

  renderContent() {
    return this.rows.map((row, index) => 
      <tr key={index}>
        {
          !this.withoutRowNum && 
            <td key={index} className={this.tableHeaderRowClass}>
              {this.renderRowNum ? this.renderRowNum(row, index) : index}
            </td>
        }
        {this.cols.map(c => {
          if (row[c.key] === undefined || row[c.key] === null) {return}
          if (row[c.key]) {
            return <td key={c.key}>{ row[c.key] }</td>
          }
        })}
      </tr>
    )
  }

	render() {
    return (
      <div className={this.className}>
        <table className={this.tableClassName}>
          <tbody>
            {this.renderHeader()}
            {this.renderContent()}
          </tbody>
        </table>
      </div>
    );
  }
}

export function ExcelRenderer(file, callback) {
    return new Promise(function(resolve, reject) {
      var reader = new FileReader();
      var rABS = !!reader.readAsBinaryString;
      reader.onload = function(e) {
        /* Parse data */
        var bstr = e.target.result;
        var wb = XLSX.read(bstr, { type: rABS ? "binary" : "array" });

        /* Get first worksheet */
        var wsname = wb.SheetNames[0];
        var ws = wb.Sheets[wsname];

        /* Convert array of arrays */
        var json = XLSX.utils.sheet_to_json(ws, { header: 1 });
        var cols = make_cols(ws["!ref"]);

        var data = { rows: json, cols: cols };

        resolve(data);
        return callback(null, data);
      };
      if (file && rABS) reader.readAsBinaryString(file);
      else reader.readAsArrayBuffer(file);
    });
  }

  function make_cols(refstr) {
    var o = [],
      C = XLSX.utils.decode_range(refstr).e.c + 1;
    for (var i = 0; i < C; ++i) {
      o[i] = { name: XLSX.utils.encode_col(i), key: i };
    }
    return o;
  }
