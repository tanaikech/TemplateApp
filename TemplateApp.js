/**
 * ### Description
 * Values from the sheet range are used in the Google Document templates.
 *
 * @param {Object} object Object including Range and Documents.
 * @param {TemplateApp~sheetRangeToDocuments} callback Callback function.
 * @return {Object} Result value.
 */
function sheetRangeToDocuments(object, callback = function () { }) {
  object.type = "document";
  return new TemplateApp(object).sheetRangeToDocuments(callback);
}

/**
 * ### Description
 * Values from the sheet range are used in the Google Slide templates.
 *
 * @param {Object} object Object including Range and Slides.
 * @param {TemplateApp~sheetRangeToSlides} callback Callback function.
 * @return {Object} Result value.
 */
function sheetRangeToSlides(object, callback = function () { }) {
  object.type = "slide";
  return new TemplateApp(object).sheetRangeToSlides(callback);
}

/**
 * ### Description
 * Manually prepared values are used in the Google Document templates.
 * In this method, the text style is not used.
 *
 * @param {Object} object Object including items and Documents.
 * @param {TemplateApp~valuesToDocuments} callback Callback function.
 * @return {Object} Result value.
 */
function valuesToDocuments(object, callback = function () { }) {
  object.type = "valueDocument";
  return new TemplateApp(object).valuesToDocuments(callback);
}

/**
 * ### Description
 * Manually prepared values are used in the Google Slide templates.
 * In this method, the text style is not used.
 *
 * @param {Object} object Object including items and Slides.
 * @param {TemplateApp~valuesToSlides} callback Callback function.
 * @return {Object} Result value.
 */
function valuesToSlides(object, callback = function () { }) {
  object.type = "valueSlide";
  return new TemplateApp(object).valuesToSlides(callback);
}

class TemplateApp {
  constructor({
    range = null,
    document = null,
    slide = null,
    items = null,
    linkForGroundColorForDoc = null,
    linkForGroundColorForSlide = null,
    useStyleOfSpreadsheet = null,
    excludeTextStyles = null,
    useImageAsPlaceholder = null,
    type,
  }) {
    this.checkParams_({ range, items, document, slide });

    this.range = range;
    this.document = document;
    this.slide = slide;
    this.items = items;
    this.header;
    this.obj;

    this.linkForGroundColorForDoc = linkForGroundColorForDoc || "#1155cc";
    this.linkForGroundColorForSlide = linkForGroundColorForSlide || "#0097a7";
    this.rowsToEachPageInSlide = (this.slide && this.slide.hasOwnProperty("rowsToEachPageInSlide")) ? this.slide.rowsToEachPageInSlide : true;
    this.useStyleOfSpreadsheet = useStyleOfSpreadsheet === null ? true : useStyleOfSpreadsheet;
    this.controlParams = ["noStyle", "image", "width", "header", "fileId"];
    this.controlFromHeader = {};
    this.convMethods_ = {};
    this.updateConvMethods_(excludeTextStyles);
    this.useImageAsPlaceholder = useImageAsPlaceholder === null ? false : useImageAsPlaceholder;
    if (type == "document" && this.useImageAsPlaceholder === true) {
      this.checkDocsAPI_();
    }
  }

  // --- Main methods

  sheetRangeToDocuments(callback) {
    if (this.range) {
      this.init_();
      if (this.document) {
        this.runDocument_(this.document, callback);
      } else {
        throw new Error("Required object 'document' is not given.");
      }
    } else {
      throw new Error("Required object 'range' is not given.");
    }
    return null;
  }

  sheetRangeToSlides(callback) {
    if (this.range) {
      this.init_();
      if (this.slide) {
        this.runSlide_(this.slide, callback);
      } else {
        throw new Error("Required object 'slide' is not given.");
      }
    } else {
      throw new Error("Required object 'range' is not given.");
    }
    return null;
  }

  valuesToDocuments(callback) {
    if (this.items) {
      if (this.document) {
        this.header = this.items[0].map(({ search }) => search);
        this.parseHeader_();
        this.runSimpleDocument_(this.document, callback);
      } else {
        throw new Error("Required object 'document' is not given.");
      }
    } else {
      throw new Error("Required object 'items' is not given.");
    }
    return null;
  }

  valuesToSlides(callback) {
    if (this.items) {
      if (this.slide) {
        this.header = this.items[0].map(({ search }) => search);
        this.parseHeader_();
        this.runSimpleSlide_(this.slide, callback);
      } else {
        throw new Error("Required object 'slide' is not given.");
      }
    } else {
      throw new Error("Required object 'items' is not given.");
    }
    return null;
  }

  // ---

  checkParams_({ range, items, document, slide }) {
    if ((!range && !items) || (range && items)) {
      throw new Error("Please set a range object or an items of array.");
    }
    if ((!document && !slide) || (document && slide)) {
      throw new Error("Please set an object of document or slide of the tampleate.");
    }
  }

  updateConvMethods_(excludeTextStyles) {
    const defaultConvMethods = {
      "isBold": "setBold",
      "isItalic": "setItalic",
      "getFontFamily": "setFontFamily",
      "isStrikethrough": "setStrikethrough",
      "isUnderline": "setUnderline",
      "getForegroundColorObject": "setForegroundColor",
      "getFontSize": "setFontSize",
      "link": "setLinkUrl",
    };
    if (
      excludeTextStyles &&
      Array.isArray(excludeTextStyles) &&
      excludeTextStyles.length != 0
    ) {
      const convMethodsFromName_ = {
        "bold": "isBold",
        "italic": "isItalic",
        "fontFamily": "getFontFamily",
        "strikethrough": "isStrikethrough",
        "underline": "isUnderline",
        "foregroundColor": "getForegroundColorObject",
        "fontSize": "getFontSize",
        "link": "link",
      };
      const tempAr = excludeTextStyles.reduce((ar, e) => {
        if (convMethodsFromName_.hasOwnProperty(e)) {
          ar.push(convMethodsFromName_[e]);
        }
        return ar;
      }, []);
      const temp = Object.entries(defaultConvMethods).filter(([k]) => !tempAr.includes(k));
      this.convMethods_ = Object.fromEntries(temp);
    } else {
      this.convMethods_ = defaultConvMethods;
    }
  }

  checkDocsAPI_() {
    try {
      const sample_ = Docs; // Dummy.
    } catch (e) {
      if (e.message == "Docs is not defined") {
        throw new Error(`IMPORTANT: When "useImageAsPlaceholder" is "true", please enable Google Docs API at Advanced Google services. After Docs API was enabled, please run the script again.`);
      }
    }
  }

  parseSheetRange_(range) {
    const sheet = range.getSheet();
    const spreadsheet = sheet.getParent();
    const spreadsheetId = spreadsheet.getId();
    const sheetName = sheet.getSheetName();
    return [spreadsheetId, sheetName];
  }

  parseHeader_() {
    const header = this.header;
    if (header.length != [...new Set(header)].length) {
      throw new Error("Same header values are existing. Please confirm the header row.");
    }
    const controlParams = this.controlParams.map(e => e.toUpperCase());
    this.controlFromHeader = header.reduce((o, h) => {
      const temp = h.replace(/{{|}}/g, "").split("_").reduce((oo, e) => {
        const t = e.split(":").map(f => {
          const tt = f.trim();
          return isNaN(tt) ? tt : Number(tt);
        });
        if (controlParams.includes(t[0].toUpperCase())) {
          oo[t[0]] = t.length == 1 ? true : t[1];
        }
        return oo;
      }, {});
      if (Object.keys(temp).length > 0) {
        o[h] = temp;
      }
      return o;
    }, {});
  }

  getImages_(spreadsheetId, sheetName) {
    const images = DocsServiceApp.openBySpreadsheetId(spreadsheetId).getSheetByName(sheetName).getImages();
    images.forEach(function (r) {
      this.obj[r.range.row - 2][r.range.col - 1].to.blob = r.image.blob;
    }, this);
    return this.obj;
  }

  init_() {
    const { header, obj } = this.parseData_(this.range);
    this.header = header;
    this.obj = obj;
    this.parseHeader_();
    if (
      this.header.some(e => e.includes("image")) &&
      this.obj.some(r => r.some(c => c.to.getImageBlob))
    ) {
      this.getImages_(...this.parseSheetRange_(this.range));
    }
  }

  // ref: https://tanaikech.github.io/2018/08/20/replacing-text-to-image-for-google-document-using-google-apps-script/
  replaceTextToImage_(body, searchText, image, width, height) {
    const next = body.findText(searchText);
    if (!next) return;
    const ele = next.getElement();
    ele.asText().setText("");
    const parent = ele.getParent();
    const parentType = parent.getType();
    let c;
    if (parentType == DocumentApp.ElementType.PARAGRAPH) {
      c = parent.asParagraph();
    } else if (parentType == DocumentApp.ElementType.LIST_ITEM) {
      c = parent.asListItem();
    } else {
      console.log(`In the current version, an image cannot be inserted to "${parentType.toString()}".`);
      return null;
    }
    const img = c.insertInlineImage(0, image);
    if (!height) {
      const w = img.getWidth();
      const h = img.getHeight();
      img.setWidth(width);
      img.setHeight(width * h / w);
    } else {
      img.setWidth(width);
      img.setHeight(height);
    }
    return next;
  }

  parseData_(range) {
    const [header, ...displayValues] = range.getDisplayValues();
    const [, ...v] = range.getValues();
    const [, ...formulas] = range.getFormulas();
    const [, ...richTextValues] = range.getRichTextValues();
    const obj = richTextValues.map((r, i) => r.map((c, j) => {
      const outputObj = { text: displayValues[i][j], getImageBlob: (typeof v[i][j] == "object" && v[i][j].toString() == "CellImage") };
      const runs = c.getRuns();
      if (runs.length > 0) {
        const exclude = ["copy", "getForegroundColor", "toString"];
        const temp = c.getRuns().map(ru => {
          const styleObj = ru.getTextStyle();
          const style = {};
          const link = ru.getLinkUrl();
          if (link) {
            style["link"] = link;
          }
          for (let f in styleObj) {
            if (!exclude.includes(f)) {
              style[f] = f == "getForegroundColorObject" ? styleObj[f]().asRgbColor().asHexString() : styleObj[f]();
            }
          }
          const start = ru.getStartIndex();
          const end = ru.getEndIndex();
          return { text: outputObj.text.slice(start, end), start, end, style };
        });
        outputObj.runs = temp;
        if (header[j].includes("image") && outputObj.text && header[j].includes("fileId")) {
          outputObj.fileId = outputObj.text;
          outputObj.getImageBlob = true;
          outputObj.blob = DriveApp.getFileById(outputObj.text).getBlob();
          const mimeType = outputObj.blob.getContentType();
          if (!mimeType.includes("image")) {
            throw new Error(`MimeType of ${mimeType} cannot be used.`);
          }
        } else if (header[j].includes("image") && outputObj.text && !header[j].includes("fileId")) {
          outputObj.link = outputObj.text;
        } else if ((header[j].includes("image") && !outputObj.text && formulas[i][j] && (/^\=IMAGE\(/i).test(formulas[i][j]))) {
          outputObj.formula = formulas[i][j];
          outputObj.link = formulas[i][j].match(/^\=IMAGE\(['"](.*)['"]/i)[1];
        }
      }
      return { "from": header[j], "to": outputObj };
    }));
    return { header, obj };
  }

  methodConversion_({ run, classObj, pos, linkForGroundColor }) {
    Object.entries(run.style).forEach(function ([k, v]) {
      if (!this.convMethods_[k] || !classObj) return;
      const args = pos ? [...pos, v] : [v];
      classObj[this.convMethods_[k]](...args);
    }, this);
    if (run.style.hasOwnProperty("link")) {
      if (run.style.isUnderline == false) {
        const args = pos ? [...pos, true] : [true];
        classObj[this.convMethods_["isUnderline"]](...args);
      }
      if (run.style.getForegroundColorObject == "#000000") {
        const args = pos ? [...pos, linkForGroundColor] : [linkForGroundColor];
        classObj[this.convMethods_["getForegroundColorObject"]](...args);
      }
    }
  }

  getImagesByDocsAPI_(doc) {
    const { inlineObjects, positionedObjects } = Docs.Documents.get(doc.getId(), { fields: "inlineObjects,positionedObjects" });
    let res = [];
    if (inlineObjects) {
      const r = Object.values(inlineObjects)
        .filter(({ inlineObjectProperties: { embeddedObject: { title } } }) => title)
        .map(({ objectId, inlineObjectProperties: { embeddedObject: { title, imageProperties: { contentUri } } } }) => [title.trim(), { objectId, contentUri }]);
      res = [...res, ...r];
    }
    if (positionedObjects) {
      const r = Object.values(positionedObjects)
        .filter(({ positionedObjectProperties: { embeddedObject: { title } } }) => title)
        .map(({ objectId, positionedObjectProperties: { embeddedObject: { title, imageProperties: { contentUri } } } }) => [title.trim(), { objectId, contentUri }]);
      res = [...res, ...r];
    }
    return res.reduce((o, [a, b]) => (o[a] = o[a] ? [...o[a], b] : [b], o), {});
  }

  replaceImageToImageByDocsAPI_({ r, doc }) {
    const obj = this.getImagesByDocsAPI_(doc);
    const inputObject = r.reduce((o, e) => {
      if (obj[e.from]) {
        if (e.to.getImageBlob) {
          obj[e.from].forEach(({ objectId }) => o[objectId] = e.to.blob);
        } else if (e.to.hasOwnProperty("link")) {
          obj[e.from].forEach(({ objectId }) => o[objectId] = e.to.link);
        }
      }
      return o;
    }, {});
    const blobs = Object.entries(inputObject).filter(([_, v]) => typeof v == "object" && v.toString() == "Blob")
    if (blobs.length > 0) {
      const tempDoc = DocumentApp.create("tempDoc by TemplateApp");
      blobs.forEach(([from, blob]) => tempDoc.getBody().appendImage(blob).setAltTitle(from));
      tempDoc.saveAndClose();
      const obj2 = this.getImagesByDocsAPI_(tempDoc);
      const obj3 = Object.entries(inputObject).reduce((o, [k, v]) => {
        if (obj2[k]) {
          obj2[k].forEach(({ contentUri }) => o[k] = contentUri);
        } else {
          o[k] = v;
        }
        return o;
      }, {});
      const tempDocFile = DriveApp.getFileById(tempDoc.getId());
      tempDocFile.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      const requests = Object.entries(obj3).map(([from, to]) => ({ replaceImage: { imageObjectId: from, uri: to } }));
      Docs.Documents.batchUpdate({ requests }, doc.getId());
      tempDocFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
      tempDocFile.setTrashed(true);
    }
  }

  replaceImageToImageBySlidesService_({ r, page }) {
    const inputObject = r.reduce((ar, e) => {
      if (e.to.getImageBlob) {
        ar.push([e.from, e.to.blob]);
      } else if (e.to.hasOwnProperty("link")) {
        ar.push([e.from, e.to.link]);
      }
      return ar;
    }, []);
    const srcObj = page.getImages().reduce((o, img) => (o[img.getTitle().trim()] = img, o), {});
    inputObject.forEach(([from, to]) => {
      if (srcObj[from]) {
        srcObj[from].replace(to);
      }
    });
  }

  runDocument_({ documents }, callback) {
    const useStyleOfSpreadsheet = this.useStyleOfSpreadsheet;
    const linkForGroundColor = this.linkForGroundColorForDoc;
    const controlFromHeader = this.controlFromHeader;
    documents.forEach(function (doc, idx) {
      callback({ status: "process", message: `Start: document ${idx + 1}` });
      const body = doc.getBody();
      const r = this.obj[idx];
      if (!r) {
        callback({ status: "warning", message: `Number of Documents is larger than that of rows. Please confirm your range and Documents again. In this case, the number of rows except for the header row is processed to the documents. In this run, the number of rows (${this.obj.length}) is processed.` });
        return;
      }
      r.forEach(function ({ from, to }) {
        if (from.includes("image")) {
          if (!to.blob && !to.link && !to.fileId) {
            throw new Error("A blob or a direct link or file ID of image file of the image were not found.");
          }
          const { width, height } = controlFromHeader[from];
          const widthValue = width || 512;
          let heightValue;
          if (height) {
            heightValue = height;
          }
          const blob = to.blob || UrlFetchApp.fetch(to.link).getBlob();
          let next;
          do {
            next = this.replaceTextToImage_(body, from, blob, widthValue, heightValue);
          } while (next);
        } else {
          let search = body.findText(from);
          while (search) {
            const ele = search.getElement();
            const start = search.getStartOffset();
            const text = ele.asText().replaceText(from, to.text).asText();
            if (useStyleOfSpreadsheet === true) {
              if (!(/noStyle/i).test(from)) {
                to.runs.forEach(function (run) {
                  const offsetStart = start + run.start;
                  const offsetEnd = start + run.end;
                  const pos = [offsetStart, offsetEnd - (offsetStart == offsetEnd ? 0 : 1)];

                  this.methodConversion_({ run, classObj: text, pos, linkForGroundColor });
                }, this);
              }
            }
            search = body.findText(from, search);
          }
        }
      }, this);
      if (this.useImageAsPlaceholder) {
        this.replaceImageToImageByDocsAPI_({ r, doc });
      }
      callback({ status: "process", message: `End: document ${idx + 1}` });
    }, this);
    return;
  }

  getObjFromShapes_(page) {
    const resObj = {}
    const shapes = page.getShapes();
    shapes.forEach(s => {
      const k = s.getText().asString().trim();
      resObj[k] = resObj[k] ? [...resObj[k], s] : [s];
    });
    const tables = page.getTables();
    tables.forEach(table => {
      for (let row = 0; row < table.getNumRows(); row++) {
        const rowObj = table.getRow(row);
        for (let col = 0; col < rowObj.getNumCells(); col++) {
          const cell = rowObj.getCell(col);
          const k = cell.getText().asString().trim();
          resObj[k] = resObj[k] ? [...resObj[k], cell] : [cell];
        }
      }
    });
    return resObj;
  }

  setDataToSlides_({ page, r, useStyleOfSpreadsheet, linkForGroundColor }) {
    const shapeObj = this.getObjFromShapes_(page);
    r.forEach(function ({ from, to }) {
      if (from.includes("image")) {
        if (!to.blob && !to.link && !to.fileId) {
          throw new Error("A blob or a direct link or file ID of image file of the image were not found.");
        }
        const sps = shapeObj[from];
        if (sps) {
          const { width, height } = this.controlFromHeader[from];
          const widthValue = (width || 512) / 1.33333;
          sps.forEach(function (sp) {
            if (sp.toString() == "Shape") {
              let heightValue;
              const top = sp.getTop();
              const left = sp.getLeft();
              const image = sp.replaceWithImage(to.blob || to.link).setTop(top).setLeft(left);
              if (!height) {
                const ratio = image.getHeight() / image.getWidth();
                heightValue = ratio * widthValue;
              } else {
                heightValue = height / 1.33333;
              }
              image.setWidth(widthValue).setHeight(heightValue);
            }
          }, this);
        }
      } else {
        const sps = shapeObj[from];
        if (sps) {
          sps.forEach(function (sp) {
            const tr = sp.getText();
            tr.replaceAllText(tr.asString(), to.text);
            if (useStyleOfSpreadsheet === true) {
              if (!(/noStyle/i).test(from)) {
                to.runs.forEach(function (run) {
                  const tss = tr.getRange(run.start, run.end).getTextStyle();
                  this.methodConversion_({ run, classObj: tss, linkForGroundColor });
                }, this);
              }
            }
          }, this);
        }
      }
      if (this.useImageAsPlaceholder) {
        this.replaceImageToImageBySlidesService_({ r, page });
      }
    }, this);
  }

  runSlide_({ slides }, callback) {
    const rowsToEachPageInSlide = this.rowsToEachPageInSlide;
    const useStyleOfSpreadsheet = this.useStyleOfSpreadsheet;
    const linkForGroundColor = this.linkForGroundColorForSlide;
    if (rowsToEachPageInSlide) {
      const slide = slides[0];
      const pages = slide.getSlides();
      this.obj.forEach(function (r, i) {
        callback({ status: "process", message: `Start: page ${i + 1} in a Google Slide` });
        const page = pages[i];
        if (!page) {
          callback({ status: "warning", message: `Number of pages in a Google Slide is smaller than that of rows. Please confirm your range and Google Slide again. In this case, the number of rows except for the header row is processed to the pages. In this run, the number of rows (${this.obj.length}) is processed.` });
          return;
        }
        this.setDataToSlides_({ page, r, useStyleOfSpreadsheet, linkForGroundColor });
        callback({ status: "process", message: `End: page ${i + 1} in a Google Slide` });
      }, this);
    } else {
      this.obj.forEach(function (r, idx) {
        callback({ status: "process", message: `Start: Google Slide ${idx + 1}` });
        const slide = slides[idx];
        if (!slide) {
          callback({ status: "warning", message: `Number of Slides is smaller than that of rows. Please confirm your range and Google Slides again. In this case, the number of rows except for the header row is processed to the pages. In this run, the number of rows (${this.obj.length}) is processed.` });
          return;
        }
        const page = slide.getSlides()[0];
        this.setDataToSlides_({ page, r, useStyleOfSpreadsheet, linkForGroundColor });
        callback({ status: "process", message: `End: Google Slide ${idx + 1}` });
      }, this);
    }
    return;
  }

  runSimpleDocument_({ documents }, callback) {
    const controlFromHeader = this.controlFromHeader;
    documents.forEach(function (doc, i) {
      callback({ status: "process", message: `Start: document ${i + 1}` });
      const body = doc.getBody();
      if (!this.items[i]) {
        callback({ status: "warning", message: `Number of rows is smaller than that of Documents. Please confirm rows and Documents. In this case, the number of Documents (${documents.length}) is processed.` });
        return;
      }
      this.items[i].forEach(function ({ search, replace }) {
        if (search === null || replace === null) {
          throw new Error("Please confirm the values of `search` and `replace`");
        }
        if (
          typeof replace == "object" &&
          replace.toString() == "Blob" &&
          replace.getContentType().includes("image")
        ) {
          const { width, height } = controlFromHeader[search];
          const widthValue = width || 512;
          let heightValue;
          if (height) {
            heightValue = height;
          }
          let next;
          do {
            next = this.replaceTextToImage_(body, search, replace, widthValue, heightValue);
          } while (next);
        } else {
          body.replaceText(search, replace);
        }
      }, this);
      callback({ status: "process", message: `End: document ${i + 1}` });
    }, this);
    return;
  }

  runSimpleSlideProcess_({ r, page }) {
    const shapeObj = this.getObjFromShapes_(page);
    r.forEach(function ({ search, replace }) {
      if (shapeObj[search]) {
        if (typeof replace == "object" && replace.toString() == "Blob" && replace.getContentType().includes("image")) {
          const width = search.match(/width:(\d*)/i);
          const widthValue = (width ? Number(width[1]) : 512) / 1.33333;
          shapeObj[search].forEach(function (sp) {
            const top = sp.getTop();
            const left = sp.getLeft();
            const image = sp.replaceWithImage(replace).setTop(top).setLeft(left);
            const ratio = image.getHeight() / image.getWidth();
            image.setWidth(widthValue).setHeight(ratio * widthValue);
          }, this);
        } else {
          shapeObj[search].forEach(function (shape) {
            const tr = shape.getText();
            tr.replaceAllText(tr.asString(), replace);
          }, this);
        }
      }
    }, this);
  }

  runSimpleSlide_({ slides }, callback) {
    const rowsToEachPageInSlide = this.rowsToEachPageInSlide;
    if (rowsToEachPageInSlide) {
      const slide = slides[0];
      const pages = slide.getSlides();
      pages.forEach(function (page, idx) {
        callback({ status: "process", message: `Start: page ${idx + 1} in a Google Slide` });
        const r = this.items[idx];
        if (!r) {
          callback({ status: "warning", message: `Number of pages of the Slide is larger than that of items. Please confirm your range and Slide again. In this case, the number of items (${this.items.length}) is processed.` });
          return;
        }
        this.runSimpleSlideProcess_({ r, page });
        callback({ status: "process", message: `End: page ${idx + 1} in a Google Slide` });
      }, this);
    } else {
      slides.forEach(function (slide, idx) {
        callback({ status: "process", message: `Start: Google Slide ${idx + 1}` });
        const page = slide.getSlides()[0];
        const r = this.items[idx];
        if (!r) {
          callback({ status: "warning", message: `Number of Slides is larger than that of items. Please confirm your range and Slides again. In this case, the number of items (${this.items.length}) is processed.` });
          return;
        }
        this.runSimpleSlideProcess_({ r, page });
        callback({ status: "process", message: `End: Google Slide ${idx + 1}` });
      }, this);
    }
    return;
  }
}
