function readEntries(reader) {
  return new Promise((resolve, reject) => {
    reader.readEntries(resolve, reject);
  });
}

function readFile(entry) {
  return new Promise((resolve, reject) => {
    entry.file(resolve, reject);
  });
}

async function readDirectory(entry, prefix = "") {
  const reader = entry.createReader();
  const allEntries = [];
  for (;;) {
    const entries = await readEntries(reader);
    if (!entries.length) break;
    allEntries.push(...entries);
  }

  const files = [];
  for (const child of allEntries) {
    files.push(...await readEntry(child, `${prefix}${entry.name}/`));
  }
  return files;
}

async function readEntry(entry, prefix = "") {
  if (entry.isFile) {
    const file = await readFile(entry);
    return [{ file, relativePath: `${prefix}${file.name}` }];
  }
  if (entry.isDirectory) {
    return readDirectory(entry, prefix);
  }
  return [];
}

export async function filesFromDataTransfer(dataTransfer) {
  const items = Array.from(dataTransfer?.items || []);
  const entries = items
    .map((item) => item.webkitGetAsEntry?.())
    .filter(Boolean);

  if (entries.length) {
    const groups = await Promise.all(entries.map((entry) => readEntry(entry)));
    return groups.flat();
  }

  return Array.from(dataTransfer?.files || []).map((file) => ({
    file,
    relativePath: file.webkitRelativePath || file.name,
  }));
}

export function filesFromInput(fileList) {
  return Array.from(fileList || []).map((file) => ({
    file,
    relativePath: file.webkitRelativePath || file.name,
  }));
}
