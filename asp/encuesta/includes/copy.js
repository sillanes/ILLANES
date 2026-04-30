async function copy(value) {
  let node;

  try {
    node = await createNode(value);
    await copyContent(node);
    resetEnv(node);
  } catch (err) {
    resetEnv(node);
  }

  return undefined;
}

function createNode(content) {
  // remove current selection
  window.getSelection().removeAllRanges();
  // create node and populate with the selectable content
  const newDiv = document.createElement('div');
  document.body.appendChild(newDiv);
  const newContent = document.createTextNode(content);
  newDiv.appendChild(newContent);
  return Promise.resolve(newDiv);
}

function copyContent(node) {
  return new Promise((resolve, reject) => {
    // select the content
    const range = document.createRange();
    range.selectNodeContents(node);
    window.getSelection().addRange(range);
    try {
      // copy the content
      const successful = document.execCommand('copy');
      const msg = successful ? 'successful' : 'unsuccessful';
      // console.log('copy command was ' + msg);
      return resolve(node);
    } catch (e) {
      // catch unsupported browsers
      console.error('error while copying range', e, range);
      return reject(node);
    }
  });
}

function resetEnv(node) {
  node.remove();
  window.getSelection().removeAllRanges();
}
