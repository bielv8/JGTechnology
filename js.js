// Seleciona o botão de envio do formulário e adiciona um ouvinte de evento para o evento "click"
document.getElementById("Cadastrar").addEventListener("click", function(){
  
    // Coleta os dados do formulário
    var nome = document.getElementById("nome").value;
    var idade = document.getElementById("idade").value;
    var email = document.getElementById("email").value;
    
    // Cria um objeto com os dados do formulário
    var pessoa = { Nome: nome, Idade: idade, Email: email };
    
    // Cria uma planilha do Excel
    var wb = XLSX.utils.book_new();
    
    // Adiciona a planilha com os dados do formulário ao arquivo do Excel
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([pessoa]), "Pessoas");
    
    // Salva o arquivo do Excel
    XLSX.writeFile(wb, 'cadastro.xlsx');
    
    // Limpa os campos do formulário após o envio
    document.getElementById("nome").value = "";
    document.getElementById("idade").value = "";
    document.getElementById("email").value = "";
  });
  