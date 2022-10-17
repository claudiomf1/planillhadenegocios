// alasql("CREATE TABLE test (language INT, hello STRING)");
// alasql("INSERT INTO test VALUES (1,'Helloooooooo!')");
// alasql("INSERT INTO test VALUES (2,'Claudio!')");
// alasql("INSERT INTO test VALUES (3,'Bonjour!')");
// console.log(alasql("SELECT * FROM test WHERE language > 1"));

let rg = "O cpf de fulado é: 124.557.895-22 (ok?)"

console.log(rg.match(/\D+/gm))

let cpfs = "680.870.540-24
140.848 .430 - 75
336.676 .580 - 10 "