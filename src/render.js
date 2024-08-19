const tb = require('./render_thingsboard');
const chuan = require('./render_chuan');

handlePayload = function(newPayload, myPayload, projectId){
  if (projectId === tb.projectId) return tb.handlePayload(newPayload, myPayload);
  else if (projectId === chuan.projectId) return chuan.handlePayload(newPayload, myPayload);
}

renderXlsx = function(sheet, payload, projectId){
  if (projectId === tb.projectId) return tb.renderXlsx(sheet, payload);
  else if (projectId === chuan.projectId) return chuan.renderXlsx(sheet, payload);
}

renderPdf = function(payload, projectId){
  if (projectId === tb.projectId) return tb.renderPdf(payload);
  else if (projectId === chuan.projectId) return chuan.renderPdf(payload);
}

module.exports = {handlePayload, renderXlsx, renderPdf};