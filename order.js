console.log("***********");

const array = [95,7, 5, 7, 9, 4, 7, 5, 2, 4, 8, 51, 2];

 let newArray = [];
for (let i = 0; i < array.length; i++) {
  for (let j = 0; j < array.length; j++) {
    const rootI = array[j]//7, 
    const brachJ = array[j+1]//7,5,7,9,4
    if(rootI>brachJ){
        array[j] = brachJ 
        array[j+1] = rootI
    }
  }
}

console.log("---");
console.log(array);


