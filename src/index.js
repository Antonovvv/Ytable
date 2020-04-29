import React from 'react';
import ReactDOM from 'react-dom';
import XLSX from 'xlsx';
import './index.css';

const utils = XLSX.utils

class App extends React.Component {
    constructor(props) {
        super(props);
        this.sheet = {};    //当前显示worksheet
        this.scope = {};      //当前显示单元格范围
        this.state = {
            tableHTML: '',
            body: [],
            head: [],
            left: [],
            contentIn: '',  //待输入内容
            cellIn: {       //目标单元格
                start: '',
                end: '',
            },
        }
    }

    handleContentInput(e) {
        this.setState({
            contentIn: e.target.value
        })
    }
    handleStartCellInput(e) {
        this.setState({
            cellIn: {...this.state.cellIn, start: e.target.value}
        })
    }
    handleEndCellInput(e) {
        this.setState({
            cellIn: {...this.state.cellIn, end: e.target.value}
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
            this.sheet = workSheet;
            this.update();
        }
        reader.readAsBinaryString(file);
    }

    replace() {
        var start = this.state.cellIn.start.toUpperCase();
        var end = this.state.cellIn.end.toUpperCase();
        if (!document.getElementById('-' + start) || !document.getElementById('-' + end)) {
            alert('单元格不存在');return;
        }
        var startCell = XLSX.utils.decode_cell(start);
        var endCell = XLSX.utils.decode_cell(end);
        if (startCell.c !== endCell.c) {
            alert('暂不支持多列'); return;
        }
        const contents = this.state.contentIn.split(' ')
        console.log(contents)

        var newSheet = {...this.sheet}
        for (let i = startCell.r; i <= endCell.r; i++) {
            var col = XLSX.utils.encode_col(startCell.c)
            var cell = newSheet[col + (i + 1)]
            if (cell) {
                cell.w = contents[i - startCell.r]
            } else {
                newSheet[col + (i + 1)] = {
                    v: contents[i - startCell.r],
                    t: 's',
                    w: contents[i - startCell.r]
                }
            }
        }
        this.sheet = newSheet
        console.log(newSheet)
        //this.updateSheet(newSheet)
    }

    update() {
        this.scope = utils.decode_range(this.sheet['!ref']);
        console.log(this.scope)
        if (this.scope.e.c < 30) this.scope.e.c = 30;
        if (this.scope.e.r < 30) this.scope.e.r = 30;
        this.table.sheetChanged();
        this.setState({
            body: sheetToTable(this.sheet, this.scope),
            head: makeCols(this.scope),
            left: makeRows(this.scope)
        })
    }

    /*updateSheet(sheet) {
        var html = XLSX.utils.sheet_to_html(sheet, {header: '', footer: ''})
        var tableHTML = htmlToTable(html)

        this.setState({
            sheetData: sheet,
        })
    }*/

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
                    <div className="controller-input">
                        <div className="form-group">
                            <label>待输入内容</label>
                            <input className="form-control" onChange={(e) => this.handleContentInput(e)}></input>
                        </div>
                        <div className="form-group">
                            <label>目标单元格</label>
                            <input className="form-control cell-input" onChange={(e) => this.handleStartCellInput(e)}/>:
                            <input className="form-control cell-input" onChange={(e) => this.handleEndCellInput(e)}/>
                        </div>
                    </div>
                    <div className="action-box">
                        <button className="btn btn-primary action-btn" onClick={() => this.replace()}>替换</button>
                        <div className="btn-group" role="group">
                            <button type="button" className="btn btn-default"><span className="glyphicon glyphicon-chevron-left"/></button>
                            <button type="button" className="btn btn-default"><span className="glyphicon glyphicon-chevron-right"/></button>
                        </div>
                    </div>
                </div>
                <OutTable body={this.state.body} head={this.state.head} left={this.state.left} scope={this.scope} onRef={(ref) => this.table = ref}/>
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

    suppress(event) {   //捕获事件
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
    constructor(props) {
        super(props);
        this.hasSheet = false;
        this.resized = false;
        this.state = {
            headerStyle: {
                width: 0,
                height: 0
            },
            cellStyle: {
                widths: [],
                heights: [],
            },
        }
    }

    componentDidMount() {
        console.log('mount')
        this.props.onRef(this);     //返回自身的反向引用
        var thead = document.getElementById('thead');
        var tleft = document.getElementById('tleft');
        var tbody = document.getElementById('tbody');
        var bodyScroll = true, headScroll = true, leftScroll = true;
        if (thead && tleft && tbody) {      //绑定滚动事件
            tbody.addEventListener('scroll', function(e) {
                if (bodyScroll) {       //由body触发
                    headScroll = false; leftScroll = false; //标识head和left的scroll事件不是由自身触发
                    thead.scrollLeft = tbody.scrollLeft;
                    tleft.scrollTop = tbody.scrollTop;
                } else {                //由head或left触发
                    bodyScroll = true;  //恢复默认状态，完成一次scroll同步
                }
            })
            thead.addEventListener('scroll', function(e) {
                if (headScroll) {
                    bodyScroll = false;
                    tbody.scrollLeft = thead.scrollLeft;
                } else {
                    headScroll = true;
                }
            })
            tleft.addEventListener('scroll', function(e) {
                if (leftScroll) {
                    bodyScroll = false;
                    tbody.scrollTop = tleft.scrollTop;
                } else {
                    leftScroll = true;
                }
            })
        }
        window.onresize = () => {
            //this.resizeTable()
        }
    }

    sheetChanged() {
        if (!this.hasSheet) this.hasSheet = true;
        this.resized = false;
    }

    componentDidUpdate() {
        if (!this.resized) this.resizeTable()
        console.log('update!')
    }

    resizeTable() {     //自动布局生成表格后根据尺寸重新渲染表头和行号列
        var source;
        var colWidths = [], rowHeights = [];
        var hSize = {};
        source = document.getElementById('hcol-');
        hSize.height = source.offsetHeight;
        hSize.width = source.offsetWidth;
        for (let i = 0; i <= this.props.scope.e.c; ++i) {
            source = document.getElementById('hcol-' + utils.encode_col(i));
            if (source) colWidths[i] = source.offsetWidth + 1;  //保证单元格不会因为宽度取整变小而高度超常
        }
        for (let i = 0; i <= this.props.scope.e.r; ++i) {
            source = document.getElementById('hrow-' + utils.encode_row(i));
            if (source) rowHeights[i] = source.offsetHeight;
        }
        this.resized = true;
        this.setState({
            headerStyle: hSize,
            cellStyle: {
                widths: colWidths,
                heights: rowHeights
            },
        })
    }

    render() {
        console.log('render')
        var headCols = this.props.head;
        var leftRows = this.props.left;
        return (
            <div style={{visibility: this.hasSheet ? '' : 'hidden'}}>
                <div className="table-header-div" id="theader">
                    <table className="table table-bordered table-header" id="table-header">
                        <thead><tr className="active"><td style={this.state.headerStyle}></td></tr></thead>
                    </table>
                </div>
                <div className="table-head-div" id="thead">
                    <table className="table table-bordered table-head" id="table-head">
                        <thead className="">
                            <tr className="active">
                                {headCols.map((c, i) => <th key={c.key} id={'col-' + c.name} 
                                style={{width: this.state.cellStyle.widths[i], height: this.state.headerStyle.height}}>{c.name}</th>)}
                                <th id="col-extra" className="pretend-bar-col"></th>
                            </tr>
                        </thead>
                    </table>
                </div>
                <div className="table-left-div" id="tleft">
                    <table className="table table-bordered table-left" id="table-left">
                        <tbody className="">
                            {leftRows.map((r, i) => <tr key={r.key} className="active">
                                <td id={'row-' + r.name} style={{height: this.state.cellStyle.heights[i], width: this.state.headerStyle.width}}>{r.name}</td>
                            </tr>)}
                            <tr><td id="row-extra" className="pretend-bar-row active"></td></tr>
                        </tbody>
                    </table>
                </div>
                <div className="table-body-div" id="tbody">
                    <table className="table table-bordered table-body" id="table-body" 
                    style={{tableLayout: this.resized ? 'fixed' : 'auto'}}>
                        <thead className={this.resized ? 'hiden-table-head' : ''}>
                            <tr>
                                <th id="hcol-"></th>
                                {headCols.map((c, i) => <th key={c.key} id={'hcol-' + c.name} 
                                style={{width: this.state.cellStyle.widths[i]}}>{c.name}</th>)}
                            </tr>
                        </thead>
                        <tbody className="table-content">
                            {this.props.body.map((row, index) => <tr key={index}>
                                <td id={'hrow-' + (index + 1)} className="hiden-table-left">{index + 1}</td>
                                {row.map((cell, i) => <td key={i} id={cell.info.id} t={cell.info.t} style={{minWidth: 50}}
                                colSpan={cell.info.colspan || ""} rowSpan={cell.info.rowspan || ""}>
                                    {cell.content}
                                </td>)}
                            </tr>)}
                        </tbody>
                    </table>
                </div>
            </div>
        )
    }
}

function sheetToTable(sheet, scope) {
    var tableData = [];
    if(!scope) scope = utils.decode_range(sheet['!ref']);
    for (var row = scope.s.r; row <= scope.e.r; ++row) {
        var merge = sheet['!merges'] || [];     //单元格合并信息
        var rowData = [];
        for (let col = scope.s.c; col <= scope.e.c; ++col) {
            var rowspan = 0, colspan = 0;
            for (let i = 0; i < merge.length; ++i) {
                if (row < merge[i].s.r || col < merge[i].s.c) continue; //当前单元格位于merge范围上或左
                if (row > merge[i].e.r || col > merge[i].e.c) continue; //当前单元格位于merge范围下或右
                if (row > merge[i].s.r || col > merge[i].s.c) {         //当前单元格位于merge范围内但非起点
                    rowspan = -1; break;    //该单元格无效
                } else {                                                //当前单元格为merge起点
                    rowspan = merge[i].e.r - merge[i].s.r + 1;
                    colspan = merge[i].e.c - merge[i].s.c + 1;
                    break;
                }
            }
            if (rowspan < 0) continue;  //当前单元格无效，继续遍历
            var addr = XLSX.utils.encode_cell({r: row, c: col});
            var cell = sheet[addr];
            var item = {info: {}, content: ''};
            item.content = ((cell && cell.v != null) && (cell.w)) || "";
            if (rowspan > 1) item.info.rowspan = rowspan;
            if (colspan > 1) item.info.colspan = colspan;
            item.info.t = (cell && cell.t) || 'z';
            item.info.id = '-' + addr;
            rowData.push(item);
        }
        tableData.push(rowData);
    }
    return tableData;
}

/*function parseHTML(html) {
    var start = html.indexOf('<table>');
    var end = html.lastIndexOf('</table>');
    var table = html.slice(start + 7, end);
    var tableRows = table.split('</tr>');
    var tableData = []
    for (let i = 0; i < tableRows.length; ++i) {
        var cellDatas = []
        var row = tableRows[i].slice(tableRows[i].indexOf('<tr>') + 4)
        var cells = row.split('</td>')
        for (let j = 0; j < cells.length; ++j) {
            var info = {}
            if (cells[j]) {
                var separateIndex = cells[j].indexOf('>')
                var cellInfo = cells[j].slice(cells[j].indexOf('<td ') + 4, separateIndex).split(' ')
                var cellContent = cells[j].slice(separateIndex + 1)
                
                for (let item of cellInfo) {
                    item = item.split('=')
                    var k = item[0]
                    var v = item[1]
                    v = v.slice(1, v.length - 1)
                    info[k] = v
                }
            }
            console.log(info)
        }
    }
}*/

/*function htmlToTable(html) {
    var start = html.indexOf('<table>');
    var end = html.lastIndexOf('</table>');
    var table = html.slice(start + 7, end);
    var tableRows = table.split('</tr>')
    for (let i = 0; i < tableRows.length - 1; ++i) {
        var row = tableRows[i].slice(tableRows[i].indexOf('<tr>') + 4)
        row = '<td id="hrow-' + (i + 1) + '" class="hiden-table-left">' + (i + 1) + '</td>' + row //增加行号
        tableRows[i] = '<tr>' + row
    }
    table = tableRows.join('</tr>')
    return table;
}*/

/* generate an array of column objects */
function makeCols(scope) {
    let o = [], C = scope.e.c + 1;
    //o[0] = {name: '', key: -1}
    for(let i = 0; i < C; ++i) o[i] = {name: XLSX.utils.encode_col(i), key: i}
	return o;
};

function makeRows(scope) {
    let o = [], R = scope.e.r + 1;
    for (let i = 0; i < R; ++i) o[i] = {name: XLSX.utils.encode_row(i), key:i}
    return o;
}

ReactDOM.render(
    <App />,
    document.getElementById('root')
)