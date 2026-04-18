// CORRECTED DELETE - List all documents in schedule/current/rows, then filter and delete
const listUrl = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents/schedule/current/rows?pageSize=1000`;

let deletedCount = 0;
let pageToken = '';

do {
  const url = pageToken ? `${listUrl}&pageToken=${pageToken}` : listUrl;
  const listResponse = await fetch(url, {
    method: 'GET',
    headers: { 'Authorization': `Bearer ${accessToken}` }
  });
  
  const listData = await listResponse.json();
  
  if (listData.documents) {
    const deletePromises = [];
    for (const doc of listData.documents) {
      // SAFETY CHECK: Only delete from schedule/current/rows collection AND matching state
      if (doc.name.includes('/schedule/current/rows/') && 
          doc.fields?.state?.stringValue === state) {
        deletePromises.push(
          fetch(`https://firestore.googleapis.com/v1/${doc.name}`, {
            method: 'DELETE',
            headers: { 'Authorization': `Bearer ${accessToken}` }
          })
        );
      }
    }
    
    if (deletePromises.length > 0) {
      await Promise.all(deletePromises);
      deletedCount += deletePromises.length;
      console.log(`Deleted ${deletePromises.length} documents (total: ${deletedCount})`);
    }
  }
  
  pageToken = listData.nextPageToken || '';
} while (pageToken);
