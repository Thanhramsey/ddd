<script>
import SheetTo from "../mixins/SheetTo";

export default {
  mixins: [SheetTo],
  data() {
    return {
      collection: null
    };
  },
  mounted() {
    this._callBack = this.updateJson;
    this.load();
  },
  methods: {
    async load() {
      const {
        utils: { sheet_to_json }
      } = await import("xlsx");
      this._sheet_to_json = sheet_to_json;
      this.loaded = true;
    },
    updateJson(workbook) {
      const ws = workbook.Sheets[this.sheetNameFinder(workbook)];
      this.collection = this._sheet_to_json(ws, this.options);
      var duoi5 = [];
      var tu5Den65 = [];
      var tu65Den8 = [];
      var tu8Den10 = [];
      //   var tenTruong = "";
      //   var khoi = "";
      console.log(ws);
      console.log(this.collection);
      for (let i = 4; i < this.collection.length; i++) {
        let item = this.collection[i];
        if (item["__EMPTY_9"] < 5 || item["__EMPTY_9"] == "-") {
          duoi5.push(item["__EMPTY_9"]);
        } else if (item["__EMPTY_9"] >= 5 && item["__EMPTY_9"] < 6.5) {
          tu5Den65.push(item["__EMPTY_9"]);
        } else if (item["__EMPTY_9"] >= 6.5 && item["__EMPTY_9"] < 8) {
          tu65Den8.push(item["__EMPTY_9"]);
        } else {
          tu8Den10.push(item["__EMPTY_9"]);
        }
      }
      //   this.collection.forEach(item => {
      // 	  if(item['__EMPTY_9']<5 || item['__EMPTY_9']=='-'){
      // 		  duoi5.push(item['__EMPTY_9']);
      // 	  }else if(item['__EMPTY_9']>=5 && item['__EMPTY_9']< 6.5){
      // 		  tu5Den65.push(item['__EMPTY_9']);
      // 	  }else if(item['__EMPTY_9']>= 6.5 && item['__EMPTY_9'] <8 ){
      // 		  tu65Den8.push(item['__EMPTY_9']);
      // 	  }else {
      // 		  tu8Den10.push(item['__EMPTY_9']);
      // 	  }
      //   })
      //   let ten = this.collection[4];
      //   tenTruong = ten["__EMPTY_5"];
      //   khoi = "Khối" + ten["__EMPTY_4"].substring(0, 2);
      //   var returnJson = [
      // 	  {"THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI":"STT","__EMPTY":"Họ và tên","__EMPTY_1":"Tên đăng nhập","__EMPTY_2":"Email","__EMPTY_3":"Ngày sinh","__EMPTY_4":"Lớp","__EMPTY_5":"Đơn vị","__EMPTY_6":"Huyện","__EMPTY_7":"Số lần tham gia thi","__EMPTY_8":"Thiết bị thi mới nhất","__EMPTY_9":"Điểm trung bình cuộc thi","__EMPTY_10":"KIỂM TRA 15P TIN 8- LẦN 3 - Điểm thi cao nhất","__EMPTY_11":"KIỂM TRA 15P TIN 8- LẦN 3 - Điểm phúc khảo","__EMPTY_12":"KIỂM TRA 15P TIN 8- LẦN 3 - Điểm trung bình","__EMPTY_13":"KIỂM TRA 15P TIN 8- LẦN 3 - Số câu trả lời đúng","__EMPTY_14":"KIỂM TRA 15P TIN 8- LẦN 3 - Thời gian","__EMPTY_15":"KIỂM TRA 15P TIN 8- LẦN 3 - Đề thi"},
      // 	  {
      // 		  "thông tin trường":tenTruong + ' '+ khoi,
      // 		  "Dưới 5:":duoi5.length +' Học sinh',
      // 		  "Từ 5 đến 6.5:":tu5Den65.length+' Học sinh',
      // 		  "Từ 6.5 đến 8:":tu65Den8.length+' Học sinh',
      // 		  "Từ 8 đến 10:":tu8Den10.length+' Học sinh'
      // 	  },
      // 	  {
      // 		  "Dưới 5:":parseFloat((duoi5.length/this.collection.length)*100).toFixed(2)+'%',
      // 		  "Từ 5 đến 6.5:":parseFloat((tu5Den65.length/this.collection.length)*100).toFixed(2)+'%',
      // 		  "Từ 6.5 đến 8:":parseFloat((tu65Den8.length/this.collection.length)*100).toFixed(2)+'%',
      // 		  "Từ 8 đến 10:":parseFloat((tu8Den10.length/this.collection.length)*100).toFixed(2)+'%'
      // 	  }
      //   ];
      var returnJson = [
        {
          "THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI":
            "Đơn vị: Trường THCS Nguyễn Văn Cừ   "
        },
        {
          "THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI":
            "Cuộc thi: KIỂM TRA 15P TIN 8- LẦN 3"
        },
        {
          "THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI":
            "Từ ngày: 22/11/2021   Đến ngày: 21/12/2021"
        },
        {
          "THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI": "STT",
          "Dưới 5": "Dưới 5",
          "Từ 5 đến 6.5:": "Từ 5 đến 6.5:",
          "Từ 6.5 đến 8:": "Từ 6.5 đến 8:",
          "Từ 8 đến 10:": "Từ 8 đến 10:"
        },
        {
          "THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI": 1,
          "Dưới 5:": duoi5.length + " Học sinh",
          "Từ 5 đến 6.5:": tu5Den65.length + " Học sinh",
          "Từ 6.5 đến 8:": tu65Den8.length + " Học sinh",
          "Từ 8 đến 10:": tu8Den10.length + " Học sinh"
        },
        {
          "THỐNG KÊ KẾT QUẢ CHI TIẾT HỌC VIÊN THAM GIA CUỘC THI": 2,
          "Dưới 5:":
            parseFloat((duoi5.length / this.collection.length) * 100).toFixed(
              2
            ) + "%",
          "Từ 5 đến 6.5:":
            parseFloat(
              (tu5Den65.length / this.collection.length) * 100
            ).toFixed(2) + "%",
          "Từ 6.5 đến 8:":
            parseFloat(
              (tu65Den8.length / this.collection.length) * 100
            ).toFixed(2) + "%",
          "Từ 8 đến 10:":
            parseFloat(
              (tu8Den10.length / this.collection.length) * 100
            ).toFixed(2) + "%"
        }
      ];
      this.$emit("parsed", returnJson);
    }
  },
  render(h) {
    if (this.$scopedSlots.default && this.loaded) {
      return h("div", [
        this.$scopedSlots.default({
          collection: this.collection
        })
      ]);
    }
    return null;
  }
};
</script>
