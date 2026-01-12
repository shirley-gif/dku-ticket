function __INIT_ONLY__setTicketSeqTo100() {
  PropertiesService.getScriptProperties().setProperty('TICKET_SEQ', '100');
}

function debugCheckTicketSeq_() {
  const v = PropertiesService.getScriptProperties().getProperty('TICKET_SEQ');
  Logger.log('TICKET_SEQ=' + v);
}

function sanityCheck_() {
  Logger.log('sanity ok');
}
