function main () {

    logseq.Editor.registerSlashCommand("Outlook Events", async () => {
        let result = await getOutlookEvents();
      await logseq.Editor.insertAtEditingCursor(result)})
  }
  
  // bootstrap
  logseq.ready(main).catch(console.error)


  async function getOutlookEvents() {
    try {
        let response = await fetch('http://localhost:5000/run_script', { method: 'GET' })
        let text = await response.text();
        return text;
    } catch(error) {
        console.error('Error:', error);
    }
}