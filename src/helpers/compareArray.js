export function compareArraysByResourceAndType(oldArray, newArray) {
  // Flattening arrays
  const flattenOldArray = oldArray.flat();
  const flattenNewArray = newArray.flat();

  // Filter Element not needed
  const filterArray = arr =>
    arr.filter(obj => obj.Recurso !== 'SUMA DE GASTOS INSIGNIFICANTES (< $0.01)' && obj.Recurso !== '');

  const filteredOldArray = filterArray(flattenOldArray);
  const filteredNewArray = filterArray(flattenNewArray);

  // Find unique elements
  const findUnique = (source, compare, isDeleted = false) => {
    return source
      .filter(
        srcItem =>
          !compare.some(
            cmpItem => srcItem.Recurso === cmpItem.Recurso && srcItem['Tipo de recurso'] === cmpItem['Tipo de recurso']
          )
      )
      .map(item => (isDeleted ? { ...item, Status: 'Deleted' } : { ...item, Status: 'New' }));
  };

  // Find new and deleted elements
  const newElements = findUnique(filteredNewArray, filteredOldArray);
  const deletedElements = findUnique(filteredOldArray, filteredNewArray, true);

  return { newElements, deletedElements };
}
