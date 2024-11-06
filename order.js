console.log("******1*****");

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

// console.log(array);

//Ejercicio 2
console.log("******2*****");

/*
  Input: l1 = [2,4,3], l2 = [5,6,4]
  Output: [7,0,8]
  Explanation: 342 + 465 = 807.
*/

function ListNode(val, next) {
  this.val = (val === undefined ? 0 : val); // Value of the node
  this.next = (next === undefined ? null : next); // Reference to the next node
}


var addTwoNumbers = function(l1 = new ListNode(0) , l2 =  new ListNode(0)) {
  let nodoListInit = new ListNode(0);
  let nodoList = nodoListInit
 
  let total = 0;
  let extra = 0
  while(l1 || l2 || extra){
    total = extra;
 
     if(l1){
       total += l1.val
       l1 = l1.next
     }
 
     if(l2){
       total += l2.val
       l2 = l2.next
     }
 
     num = total % 10
     extra = Math.floor(total / 10)
     nodoListInit.next = new ListNode(num)
     nodoListInit =  nodoListInit.next
  }
 
  return nodoList.next
 };

const l1 = [2,4,3]
const node1 = new ListNode(2);
const node2 = new ListNode(4);
const node3 = new ListNode(3);

node1.next = node2
 const l2 = [5,6,4]
console.log(addTwoNumbers(l1,l2));


///ejercicio 3:
var lengthOfLongestSubstring = function(s) {
  let left = 0; // Puntero izquierdo para el inicio de la ventana deslizante
  let maxLength = 0; // Almacena la longitud máxima de la subcadena sin caracteres repetidos
  let charSet = new Set(); // Set para almacenar los caracteres únicos de la subcadena actual

  // Recorremos la cadena usando el puntero derecho
  for (let right = 0; right < s.length; right++) {
      // Si el carácter en la posición 'right' ya está en el Set (lo que indica un carácter repetido),
      // movemos el puntero izquierdo 'left' hacia adelante y eliminamos los caracteres del Set
      // hasta que el carácter repetido ya no esté en el Set
      while (charSet.has(s[right])) {
          charSet.delete(s[left]); // Eliminamos el carácter en la posición 'left' del Set
          left++; // Movemos el puntero izquierdo a la derecha
      }

      // Agregamos el carácter en la posición 'right' al Set
      charSet.add(s[right]);//abc
      
      // Actualizamos la longitud máxima de la subcadena sin repetidos
      // Calculamos la longitud de la ventana actual como 'right - left + 1'
      maxLength = Math.max(maxLength, right - left + 1);
  }

  return maxLength; // Devolvemos la longitud de la subcadena sin caracteres repetidos más larga
};
