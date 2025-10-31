Feature: Criar fazenda
  
  #Evidência: https://drive.google.com/file/d/.......
  #Resultado: SUCESSO✅
  Scenario: 01) Cadastrar Primeira Fazenda
    Given que o usuário está logado e com conexão ativa
    When ele clica no botão "Cadastrar Primeira Fazenda"
    And Preencher todos os campos com informações válidas
    Then uma fazenda deve ser criada com sucesso

  #Evidência: https://drive.google.com/file/d/1...............
  #Resultado: SUCESSO✅
  Scenario: 02) Editar Fazenda
    Given que o usuário está logado e com uma fazenda já criada
    When ele clica na fazenda
    Then deve ser possível editar os dados da fazenda
