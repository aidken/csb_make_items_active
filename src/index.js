import "./styles.css";
import XLSX from "xlsx";
import ejs from "ejs";

let template1 = `Description = Make Items Active
[wait app]
[wait inp inh]

;;; go to RVKOE5
"rvkoe5
[enter]
[wait inp inh]
wait 10 sec until FieldAttribute 0000 at (4,50)
wait 10 sec until cursor at (4,51)
[wait app]
[wait inp inh]

;;; start program
"53
[enter]
[wait inp inh]
wait 10 sec until FieldAttribute 0000 at (9,1)
[wait app]
[wait inp inh]
`;

let template2 = `
;;; item <%= item %>
[pf6]
[wait inp inh]
wait 10 sec until FieldAttribute 0000 at (7,23)
wait 10 sec until cursor at (7,24)
[wait app]
[wait inp inh]
[tab field]
[tab field]
[tab field]
[tab field]
"710536
[tab field]
"<%= item %>
[field exit]
"<%= dateFrom %>
[field exit]
"<%= dateTo %>
[field exit]
[enter]
[wait inp inh]
wait 10 sec until FieldAttribute 0000 at (7,23)
[wait app]
[wait inp inh]
[pf6]
[wait inp inh]
wait 10 sec until FieldAttribute 0000 at (9,1)
[wait app]
[wait inp inh]
`;

let dateFrom = (function () {
  let x = new Date();
  let dd = String(x.getDate()).padStart(2, "0");
  let mm = String(x.getMonth() + 1).padStart(2, "0"); //January is 0!
  let yy = String(x.getFullYear() - 2000);
  return yy + mm + dd;
})();

let dateTo = (function () {
  let x = new Date();
  x.setDate(x.getDate() + 10);
  let dd = String(x.getDate()).padStart(2, "0");
  let mm = String(x.getMonth() + 1).padStart(2, "0"); //January is 0!
  let yy = String(x.getFullYear() - 2000);
  return yy + mm + dd;
})();

let dateFromJpn = (function () {
  let x = new Date();
  let dd = String(x.getDate());
  let mm = String(x.getMonth() + 1); //January is 0!
  let yy = String(x.getFullYear());
  return yy + '年' + mm + '月' + dd + '日';
})();

let dateToJpn = (function () {
  let x = new Date();
  x.setDate(x.getDate() + 10);
  let dd = String(x.getDate());
  let mm = String(x.getMonth() + 1); //January is 0!
  let yy = String(x.getFullYear());
  return yy + '年' + mm + '月' + dd + '日';
})();

let fileName = (function () {
  let x = new Date();
  let dd = String(x.getDate()).padStart(2, "0");
  let mm = String(x.getMonth() + 1).padStart(2, "0"); //January is 0!
  let yy = String(x.getFullYear());
  return "make_active_" + yy + "-" + mm + "-" + dd + ".mac";
})();

let ExcelToJSON = function () {
  // initialize
  let macro = template1;

  // // order class
  // let Order = function (row) {
  //   this.itemNumber = row.B.toString();
  //   this.status = row.D.toString().toLowerCase().trim();
  // };

  this.parseExcel = function (file) {
    var reader = new FileReader();

    reader.onload = function (evt) {
      let data = evt.target.result;
      let workbook = XLSX.read(data, { type: "binary" });
      let worksheet = workbook.Sheets["提出用"];
      let XL_row_object = XLSX.utils.sheet_to_json(worksheet, { header: "A" });

      XL_row_object.forEach(function (i) {
        if (
          typeof i.B !== "undefined" &&
          i.B.toString() !== "商品ｺｰﾄﾞ" &&
          i.B.toString() !== "総計"
        ) {
          if (i.D.toString().toLowerCase().trim() === "dis") {
            macro += ejs.render(template2, {
              item: i.B.toString().toLowerCase().trim(),
              dateFrom: dateFrom,
              dateTo: dateTo
            });
          } // end if
        } // end if
      }); // end XL_row_object.forEach

      // replace /n with /r/n
      macro = macro.replace(/\n/g, "\r\n");

      var pom = document.createElement("a");
      pom.setAttribute(
        "href",
        "data:text/plain;charset=utf-8," + encodeURIComponent(macro)
      );

      pom.setAttribute("download", fileName);

      pom.style.display = "none";
      document.body.appendChild(pom);

      pom.click();

      document.body.removeChild(pom);
    }; // end onload

    reader.onerror = function (ex) {
      console.log(ex);
    }; // end reader.onerror

    reader.readAsBinaryString(file);
  }; // end parseExcel
}; // end function ExcelToJSON

function handleFileSelect1(evt) {
  let files = evt.target.files; // FileList object
  let xl2json = new ExcelToJSON();
  xl2json.parseExcel(files[0]);
} // end function handleFile1

document.getElementById("app").innerHTML = `
<h2 class="jumbotron text-center" style="margin-bottom:0">Dis の品目を、これから10日間だけ Act にするマクロを作る</h2>

  <div class="container" style="margin-top:20px">

    <div class="row">

      <div class="col-sm-3">
        <ul>
          <li><a href="https://aidken.gitbook.io/fx/" target="_blank" rel="noopener noreferrer">FX の在庫が足りているかどうか確認</a>.</li>
          <li><a href="https://aidken.gitbook.io/bpcs/" target="_blank" rel="noopener noreferrer">楽天注文を BPCS へ入力するマクロスクリプトを作る</a></li>
          <li>Dis の品目を、これから10日間だけ Act にするマクロを作る</li>
        </ul>
      </div>

      <div class="col-sm-9">
      	<form enctype="multipart/form-data">
      		<p>
		        <label for='upload1'>楽天の注文の Excel ファイルをアップロードして下さい。</label>
		        <input id="upload1" type=file name="files1[]" accept='.xlsm, .xlsx'>
          </p>
          <h4>使い方</h4>
          <p>このウェブアプリは、楽天の注文のファイルをアップロードすると、Dis になっている品目を Act に切り替えるためのマクロを作成してダウンロードします。</p>
          <p>アップロードすると、エクセルファイルの中の「提出」シートを読んで、D列「Status」が「Dis」になっている品目を見つけ出します。</p>
          <p>これを、Act に切り替えるマクロを作ります。きょう ${dateFromJpn} から、${dateToJpn} までの10日間 Act に設定します。</p>
          <p>マクロのファイルは、ほんの数秒で作られると思います。マクロのファイルが出来上がったら、自動的にダウンロードされます。</p>
          <h4>履歴</h4>
          <ul>
            <li>2020年12月26日　最初のバージョンを作成。ただし出来上がったマクロファイルがきちんと動作するかどうかは未確認。</li>
            <li>2020年12月27日　このウェブアプリで作ったマクロのファイルがきちんと動作することが確認された。</li>
          </ul>
	      </form>
      </div>

    </div>

  </div>
`;

document
  .getElementById("upload1")
  .addEventListener("change", handleFileSelect1, false);
