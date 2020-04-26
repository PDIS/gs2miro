const zip = new require('node-zip')();
const JSONbig = require('json-bigint');
const fs = require("fs");
const parse = require('csv-parse/lib/sync');
const fetch = require('node-fetch')
const XLSX = require('xlsx');

module.exports = {
  gs2mrio: async (url, filename) => {
    //initial board and widgets
    let board = JSONbig.parse(`{"name":"","description":"","isPublic":false,"iconResourceId":0,"id":-6642106231252299288}`)
    let meta = JSONbig.parse(`{"version":"1.2"}`)
    let resources = JSONbig.parse(`{"resources":[]}`)
    let canvas = JSONbig.parse(`{"id":0,"widgets":{"order":[3074457347270829190,3074457347270829205,3074457347270829206,3074457347270829207,3074457347270829208,3074457347270829209,3074457347270829210,3074457347270829211,3074457347270829212,3074457347270829213,3074457347270829214,3074457347270829307,3074457347270829308],"objects":{"cliparts":[],"curves":[],"documents":[],"images":[],"lines":[],"mockups":[],"shapes":[{"x":3789.0232508362405,"y":1862.2244022886816,"rotation":0.0,"width":9381.784356255657,"height":5105.0184346357655,"style":{"st":3,"ffn":10,"fs":162,"b":0,"i":0,"u":0,"s":0,"bc":16777215,"bo":1,"brc":1710618,"bro":1,"brw":2,"brs":2,"tc":1710618,"ta":"c","tav":"m","tsc":1,"bsc":0,"VER":2},"text":"","type":"3","id":3074457347270829190},{"x":-112.79573179508407,"y":921.8032629043437,"rotation":0.0,"width":221.5439344663691,"height":1521.2485654879792,"style":{"st":3,"ffn":10,"fs":72,"b":0,"i":0,"u":0,"s":0,"bc":-1,"bo":1,"brc":1710618,"bro":1,"brw":2,"brs":2,"tc":1710618,"ta":"c","tav":"m","tsc":1,"bsc":1,"VER":2},"text":"","type":"3","id":3074457347270829206},{"x":364.5494398926239,"y":-103.64895072693434,"rotation":0.0,"width":1176.83221434994,"height":247.7041576171631,"style":{"st":3,"ffn":10,"fs":33,"b":0,"i":0,"u":0,"s":0,"bc":-1,"bo":1,"brc":1710618,"bro":1,"brw":2,"brs":2,"tc":1710618,"ta":"c","tav":"m","tsc":1,"bsc":0,"VER":2},"text":"","type":"3","id":3074457347270829308}],"stickers":[{"x":-112.79573179508407,"y":527.8661391369578,"width":199.0,"height":228.0,"scale":0.8508944565406553,"style":{"sbc":16775601,"ffn":0,"fs":64,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e問題\u003c/p\u003e\u003cp\u003e細節\u003c/p\u003e","id":3074457347270829207},{"x":-112.79573179508407,"y":710.9829710061622,"width":199.0,"height":228.0,"scale":0.8318259233900455,"style":{"sbc":13689210,"ffn":0,"fs":64,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e建議\u003c/p\u003e\u003cp\u003e解法\u003c/p\u003e","id":3074457347270829208},{"x":-112.79573179508407,"y":1058.7713986332074,"width":199.0,"height":228.0,"scale":0.7691023531853076,"style":{"sbc":16751944,"ffn":0,"fs":46,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e政府\u003c/p\u003e\u003cp\u003e回應\u003c/p\u003e\u003cp\u003e做法\u003c/p\u003e","id":3074457347270829209},{"x":-112.79573179508407,"y":888.452428321355,"width":199.0,"height":228.0,"scale":0.8013560832598544,"style":{"sbc":10931445,"ffn":0,"fs":64,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e背景\u003c/p\u003e\u003cp\u003e資料\u003c/p\u003e","id":3074457347270829210},{"x":-112.79573179508407,"y":1225.4134437165594,"width":199.0,"height":228.0,"scale":0.7691023531853076,"style":{"sbc":15821951,"ffn":0,"fs":38,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e政府回應\u003c/p\u003e\u003cp\u003e困難和\u003c/p\u003e\u003cp\u003e限制\u003c/p\u003e","id":3074457347270829211},{"x":-112.79573179508407,"y":1576.1241167690605,"width":199.0,"height":228.0,"scale":0.7691023531853084,"style":{"sbc":16448250,"ffn":0,"fs":64,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e補充\u003c/p\u003e\u003cp\u003e說明\u003c/p\u003e","id":3074457347270829212},{"x":-112.79573179508407,"y":344.5327649339815,"width":199.0,"height":228.0,"scale":0.8508944565406553,"style":{"sbc":16109864,"ffn":0,"fs":64,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e問題\u003c/p\u003e\u003cp\u003e分類 \u003c/p\u003e","id":3074457347270829213},{"x":-112.79573179508407,"y":1400.76878024281,"width":199.0,"height":228.0,"scale":0.7691023531853076,"style":{"sbc":11764923,"ffn":0,"fs":50,"fsa":1,"ta":"c","tav":"m","taw":155,"tah":156,"lh":1.1},"text":"\u003cp\u003e待釐清\u003c/p\u003e\u003cp\u003e的問題\u003c/p\u003e","id":3074457347270829214}],"texts":[{"x":-92.94876286676117,"y":205.90783940947495,"rotation":0.0,"scale":2.484936624951172,"width":83.78219999999999,"height":36.0,"style":{"st":14,"bc":-1,"bo":1,"brc":-1,"bro":1,"brw":2,"brs":2,"bsc":1,"ta":"l","tc":1710618,"tsc":1,"ffn":10,"b":0,"i":0,"u":0,"s":0,"fw":1},"text":"\u003cp\u003e顏色標記\u003c/p\u003e","id":3074457347270829205},{"x":364.5494398926239,"y":-103.64895072693434,"rotation":0.0,"scale":6.835324309724425,"width":165.83479999999997,"height":48.0,"style":{"st":14,"bc":-1,"bo":1,"brc":-1,"bro":1,"brw":2,"brs":2,"bsc":1,"ta":"c","tc":1710618,"tsc":1,"ffn":0,"b":0,"i":0,"u":0,"s":0,"fw":1},"text":"\u003cp\u003e議題脈絡分類釐清\u003c/p\u003e\u003cp\u003eIssue-based mind mapping\u003c/p\u003e","id":3074457347270829307}],"videos":[],"webScreenshots":[],"linkPreviews":[],"embeds":[],"jiraCards":[],"rallyWidgets":[],"customWidgets":[]}},"comments":[{"positionX":-49,"positionY":-24,"creatorId":3074457345878814619,"state":4,"linkerid":3074457347270829213,"messages":[{"time":1583576108000,"senderId":3074457345878814619,"text":":money_mouth:","mentions":[]}],"inspections":{"3074457345878814619":1583576202003},"isResolved":true,"id":3074457347322242179}],"links":[],"groups":[],"camera":{"a":{"x":0.0,"y":0.0},"b":{"x":0.0,"y":0.0}},"presentationVisible":true,"presentation":[],"labels":[],"emojis":{},"widgetsAliases":[]}`)

    let widgets = canvas.widgets
    let order = widgets.order
    let stickers = widgets.objects.stickers
    let lines = widgets.objects.lines
    let labels = canvas.labels

    let labelcolors = [9769132, 4279218, 15877926, 8421504, 16434960, 1232340]
    let counter = {
      'lv1': 1,
      'lv2': 1,
      'lv3': 1,
      'lv4': 1,
      'lv5': 1,
      'color': 0
    }

    const RandomID = () => {
      return Math.floor(Math.random() * 8999999999999) + 1000000000000
    }

    const MakeSticker = (id, text, color, level, counter) => {
      let n = Math.ceil(Math.sqrt(text.length))
      let textarray = []
      for (let i = 0; i < n; i++) {
        let temp = text.slice(i * n, (i + 1) * n)
        if (temp !== '') {
          textarray.push(text.slice(i * n, (i + 1) * n))
        }
      }
      let newtext = ''
      for (let i = 0; i < textarray.length; i++) {
        newtext += `<p>${textarray[i]}</p>`
      }
      let sticker = {
        "x": counter * 1000,
        "y": level * 500,
        "width": 199.0,
        "height": 228.0,
        "scale": 1,
        "style": {
          "sbc": color,
          "ffn": 0,
          "fsa": 1,
          "ta": "c",
          "tav": "m",
          "taw": 155,
          "tah": 156,
          "lh": 1.1
        },
        "text": newtext,
        "id": id
      }
      stickers.push(sticker)
      order.push(id)
    }

    const MakeLine = (from, to) => {
      let lineid = RandomID()
      let line = {
        "waypoints": [],
        "style": {
          "lc": 0,
          "ls": 2,
          "t": 1,
          "lt": 0,
          "a_start": 0,
          "a_end": 1,
          "VER": 2
        },
        "primaryLink": {
          "widgetId": from,
          "position": {
            "x": 0.5,
            "y": 1.0
          },
          "positionType": "BORDER"
        },
        "secondaryLink": {
          "widgetId": to,
          "position": {
            "x": 0.5,
            "y": 0.0
          },
          "positionType": "BORDER"
        },
        "id": lineid
      }
      lines.push(line)
      order.push(lineid)
    }

    const MakeLabel = (sticker, text) => {
      let id = RandomID()
      if (counter.color > 5) {
        counter.color = 0
      }
      let color = labelcolors[counter.color]
      let reg = /、/g
      if (reg.test(text)) {
        let textarray = text.split('、')
        textarray.map(t => {
          color = labelcolors[counter.color]
          let label = {
            "text": t,
            "style": {
              "c": color
            },
            "widgets": [sticker],
            "id": id
          }
          labels.push(label)
          order.push(id)
          counter.color++
        })
      } else {
        let label = {
          "text": text,
          "style": {
            "c": color
          },
          "widgets": [sticker],
          "id": id
        }
        labels.push(label)
        order.push(id)
        counter.color++
      }
    }

    board.name = filename
    board.id = RandomID()

    const res = await fetch(url)
    if (!res.ok) throw new Error("fetch failed");
    const buffer = await res.arrayBuffer();
    const data = new Uint8Array(buffer);
    const workbook = XLSX.read(data, {
      type: "array"
    });
    const sheets = workbook.SheetNames
    sheets.map(sheet => {
      let csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheet])
      let level = 0
      let isData = false

      let department = {
        'id': 0,
        'text': ''
      }

      let perspective = {
        'id': 0,
        'text': ''
      }

      let description = {
        'id': 0,
        'text': ''
      }

      let solution = {
        'id': 0,
        'text': ''
      }

      let response = {
        'id': 0,
        'text': ''
      }

      let difficulty = {
        'id': 0,
        'text': ''
      }

      parser = parse(csv, {
        auto_parse: true
      });

      parser.map(obj => {
        if (isData == true) {
          for (let i = 0; i < obj.length; i++) {
            let text = obj[i]
            let color = 0
            let id = RandomID()
            switch (i) {
              case 0: //填寫部會
                if (text !== '') {
                  department.id = id
                  department.text = text
                }
                break
              case 1: //問題面向
                if (text !== '') {
                  perspective.id = id
                  perspective.text = text
                  color = 16109864
                  level = 1
                  MakeSticker(id, text, color, level, counter.lv1)
                  counter.lv1++
                }
                break
              case 2: //問題細節
                if (text !== '') {
                  level = 2
                  color = 16775601
                  description.id = id
                  description.text = text
                  MakeSticker(id, text, color, level, counter.lv2)
                  counter.lv2++
                  MakeLine(perspective.id, description.id)
                }
                break
              case 3: //填寫細節的人
                if (text !== '') {
                  let isOld = false
                  labels.map(l => {
                    if (l.text === text) {
                      isOld = true
                      l.widgets.push(description.id)
                    }
                  })
                  if (isOld == false) {
                    MakeLabel(description.id, text)
                  }
                }
                break
              case 5: //解法
                if (text !== '') {
                  level = 3
                  color = 13689210
                  solution.id = id
                  solution.text = text
                  MakeSticker(id, text, color, level, counter.lv3)
                  counter.lv3++
                  MakeLine(description.id, solution.id)
                }
                break
              case 7: //填寫解法的人
                if (text !== '') {
                  let isOld = false
                  labels.map(l => {
                    if (l.text === text) {
                      isOld = true
                      l.widgets.push(solution.id)
                    }
                  })
                  if (isOld == false) {
                    MakeLabel(solution.id, text)
                  }
                }
                break
              case 9: //政府回應
                if (text !== '') {
                  level = 4
                  color = 16751944
                  response.id = id
                  response.text = text
                  MakeSticker(id, text, color, level, counter.lv4)
                  counter.lv4++
                  MakeLine(solution.id, response.id)
                  let isOld = false
                  labels.map(l => {
                    if (l.text === department.text) {
                      isOld = true
                      l.widgets.push(response.id)
                    }
                  })
                  if (isOld == false) {
                    MakeLabel(response.id, department.text)
                  }
                }
                break
              case 10: //困難
                if (text !== '') {
                  level = 5
                  color = 15821951
                  difficulty.id = id
                  difficulty.text = text
                  MakeSticker(id, text, color, level, counter.lv5)
                  counter.lv5++
                  MakeLine(response.id, difficulty.id)
                  let isOld = false
                  labels.map(l => {
                    if (l.text === department.text) {
                      isOld = true
                      l.widgets.push(difficulty.id)
                    }
                  })
                  if (isOld == false) {
                    MakeLabel(difficulty.id, department.text)
                  }
                }
                break
            }
          }
        }
        if (obj[12] === '希望協作會議達到什麼目標、產出什麼成果') {
          isData = true
        } else if (obj[0] == 1) {
          isData = false
        }
      })
    })
    zip.file('board.json', JSONbig.stringify(board))
    zip.file('meta.json', JSONbig.stringify(meta))
    zip.file('resources.json', JSONbig.stringify(resources))
    zip.file('canvas.json', JSONbig.stringify(canvas))
    const genzip = zip.generate({
      base64: false,
      compression: 'DEFLATE'
    });
    fs.writeFileSync(`${filename}.rtb`, genzip, 'binary');
  },
  deletefile: (filepath) => {
    fs.unlinkSync(filepath)
  }
};