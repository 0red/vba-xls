xls_column_name=(nr)=>{  // first column 0 as JS Array not like VBA (1)
    let w=String.fromCharCode(((nr)%26)+65)
    return (nr>26-1)? xls_column_name(Math.floor(nr/26)-1)+w:w
};
console.log(xls_column_name(25))    // "Z"
console.log(xls_column_name(26))    // "AA"
console.log(xls_column_name(26*27)) // "AAA" (702)


a=[]
for (let i=0;i<=26*27+1;i++) a.push(xls_column_name(i))
console.log(a)



xls_column_number=(str)=>str.split("").reverse().reduce((pr,cu,i)=>pr+(cu.charCodeAt(0)-64)*26**i,0)-1    

console.log(xls_column_number(xls_column_name(25)))     // "Z"
console.log(xls_column_number(xls_column_name(26)))     // "AA"
console.log(xls_column_number(xls_column_name(26*27)))  // "AAA" (702)
