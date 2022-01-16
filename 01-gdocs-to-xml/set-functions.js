function setDifference(a, b) {

  let diff = new Set(a);

  for (let elem of b) {
      diff.delete(elem)
  }

  return diff;
}
