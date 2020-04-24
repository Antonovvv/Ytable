import React from 'react';
import ReactDOM from 'react-dom';
import XLSX from 'xlsx';
import './index.css';

class App extends React.Component {
    constructor(props) {
        super(props);
        this.state = {
            data: '',
            headCols: [],   //表头列号
            leftRows: [],   //行号
            cellIn: '',
            rowCount: 0,
        }
    }

    handleCellInput(e) {
        this.setState({
            cellIn: e.target.value
        })
    }

    handleFile(file) {
        const reader = new FileReader()
        reader.onload = (e) => {
            const f = e.target.result
            var workBook = XLSX.read(f, {type: 'binary'})
            console.log(workBook)

            var workSheetName = workBook.SheetNames[0]
            var workSheet = workBook.Sheets[workSheetName]
            //var data = XLSX.utils.sheet_to_json(workSheet, {header: 1})
            var htmlData = XLSX.utils.sheet_to_html(workSheet, {header: '', footer: ''})
            var tableHTML = htmlToTable(htmlData)
            
            this.setState({
                data: tableHTML,
                headCols: makeCols(workSheet['!ref']),
                leftRows: makeRows(workSheet['!ref']),
            })
        }
        reader.readAsBinaryString(file);
    }

    test() {
        var firstRows = [];
        firstRows[0] = document.getElementById('sjs-' + this.state.cellIn.toUpperCase())
        console.log(firstRows[0] ? firstRows[0] : 'not found')
        var cellTarget = document.getElementById('col-A')
        if (cellTarget && firstRows[0]) cellTarget.style.width = firstRows[0].offsetWidth + 'px';
    }

    render() {
        return (
            <div>
                <nav className="navbar navbar-default">
                    <div className="container-fluid">
                        <div className="navbar-header">
                            <a className="navbar-brand" href="/#">歪歪</a>
                            {/*<button type="button" className="btn btn-primary navbar-btn" data-toggle="collapse" data-target="#controller"
                            aria-expanded="false" aria-controls="controller">填表</button>*/}
                        </div>
                    </div>
                </nav>
                <div className="controller" id="controller">
                    <DropFile handleFile={(f) => this.handleFile(f)}/>
                    <div className="form-group">
                        <label>单元格</label>
                        <input className="form-control cell-input" id="cell-input" onChange={(e) => this.handleCellInput(e)}/>
                    </div>
                    <button className="btn btn-primary test" onClick={() => this.test()}>查询</button>
                </div>
                <OutTable data={this.state.data} headCols={this.state.headCols} leftRows={this.state.leftRows}/>                
            </div>
        )
    }
}

class DropFile extends React.Component {
    handleDrop(event) {
        this.suppress(event)
        const files = event.dataTransfer.files
        if (files && files[0]) this.props.handleFile(files[0])
    }

    suppress(event) {
        event.stopPropagation();
        event.preventDefault();
    }

    render() {
        return (
            <div onDrop={(e) => this.handleDrop(e)} onDragEnter={this.suppress} onDragOver={this.suppress}>
                <div className="file-input">拖入文件</div>
            </div>
        )
    }
}

class OutTable extends React.Component {
    componentDidMount() {
        window.onresize = () => {
            var colWidths = this.resizeHead()
            this.resizeLeft()
            this.resizeBody(colWidths)
        }
    }

    componentDidUpdate() {  //绑定滚动事件
        var thead = document.getElementById('thead');
        var tleft = document.getElementById('tleft');
        var tbody = document.getElementById('tbody');
        if (thead && tleft && tbody) {
            tbody.addEventListener('scroll', function(e) {
                thead.scrollLeft = tbody.scrollLeft;
                tleft.scrollTop = tbody.scrollTop;
            })
        }
        var colWidths = this.resizeHead()
        this.resizeLeft()
        this.resizeBody(colWidths)
    }

    resizeHead() {
        var targetBody = document.getElementById('table-body')
        if(targetBody) targetBody.style.tableLayout = 'auto'   //令表格内容自动排列

        var sourceCell; var targetCell;
        var cols = this.props.headCols;
        var colWidths = []
        for (let i = 0; i < cols.length; i++) {
            sourceCell = document.getElementById('hcol-' + cols[i].name)
            targetCell = document.getElementById('col-' + cols[i].name)
            if (sourceCell && targetCell) {
                colWidths[i] = sourceCell.offsetWidth
                targetCell.style.width = sourceCell.offsetWidth + 'px'
            }
        }
        return colWidths
    }

    resizeLeft() {
        var targetBody = document.getElementById('table-body')
        if(targetBody) targetBody.style.tableLayout = 'auto'   //令表格内容自动排列

        var sourceCell; var targetCell;
        var rows = this.props.leftRows;
        var rowHeights = []
        for (let i = 0; i < rows.length; i++) {
            sourceCell = document.getElementById('hrow-' + rows[i].name)
            targetCell = document.getElementById('row-' + rows[i].name)
            if (sourceCell && targetCell) {
                rowHeights[i] = sourceCell.offsetHeight
                targetCell.style.height = sourceCell.offsetHeight + 'px'
            }
        }
        return rowHeights
    }

    resizeBody(widths) {    //跟据已调整的列宽度(整数)调整表内容
        var targetBody = document.getElementById('table-body')
        if(targetBody) {
            targetBody.style.right = widths[0] + 'px'
            targetBody.style.tableLayout = 'fixed'   //令表格内容宽度可设置
        }
        var targetCell;
        var cols = this.props.headCols;
        for (let i = 0; i < cols.length; i++) {
            targetCell = document.getElementById('hcol-' + cols[i].name)
            if (targetCell) {
                targetCell.style.width = widths[i] + 'px'
            }
        }
    }

    render() {
        return (
            <div className="">
                <div className="table-head-div" id="thead">
                    <table className="table table-bordered table-head" id="table-head">
                        <thead className="">
                            <tr className="active">
                                {this.props.headCols.map((c) => <th key={c.key} id={'col-' + c.name}>{c.name}</th>)}
                                <th id={'col-extra'} className="fakeScrollBar">{this.props.data ? '' : '没有表~'}</th>
                            </tr>
                        </thead>
                    </table>
                </div>
                <div className="table-left-div" id="tleft">
                    <table className="table table-bordered table-left" id="table-left">
                        <tbody className="">
                            {this.props.leftRows.map((r) => 
                            <tr key={r.key} className="active">
                                <td id={'row-' + r.name}>{r.name}</td>
                            </tr>)}
                        </tbody>
                    </table>
                </div>
                <div className="table-body-div" id="tbody">
                    <table className="table table-bordered table-body" id="table-body">
                        <thead className="hiden-table-head">
                            <tr>
                                {this.props.headCols.map((c) => <th key={c.key} id={'hcol-' + c.name}>{c.name}</th>)}
                            </tr>
                        </thead>
                        <tbody className="table-content" dangerouslySetInnerHTML={{__html: this.props.data}}></tbody>
                    </table>
                </div>
            </div>
        )
    }
}

function htmlToTable(html) {
    var start = html.indexOf('<table>');
    var end = html.lastIndexOf('</table>');
    var table = html.slice(start + 7, end);
    var tableRows = table.split('</tr>')
    for (let i = 0; i < tableRows.length - 1; i++) {
        var row = tableRows[i].slice(tableRows[i].indexOf('<tr>') + 4)
        row = '<td id="hrow-' + (i + 1) + '" class="hiden-table-left">' + (i + 1) + '</td>' + row //增加行号
        tableRows[i] = '<tr>' + row
    }
    return tableRows.join('</tr>');
}

/* generate an array of column objects */
function makeCols(ref) {
    let o = [], C = XLSX.utils.decode_range(ref).e.c + 1;
    o[0] = {name: '', key: -1}
    for(let i = 0; i < C; i++) o[i + 1] = {name: XLSX.utils.encode_col(i), key: i}
	return o;
};

function makeRows(ref) {
    let o = [], R = XLSX.utils.decode_range(ref).e.r + 1;
    for (let i = 0; i < R; i++) o[i] = {name: XLSX.utils.encode_row(i), key:i}
    return o;
}

ReactDOM.render(
    <App />,
    document.getElementById('root')
)