# Gerenciador de Unidades de Rede

## Sobre o projeto

Script desenvolvido em **PowerShell** para facilitar o gerenciamento de unidades de rede no Windows, permitindo mapear e desmapear caminhos de rede por meio de uma interface interativa no terminal.

O script automatiza tarefas que normalmente seriam realizadas manualmente pelo Explorador de Arquivos ou por comandos do Windows, oferecendo também a possibilidade de criar e remover atalhos das unidades mapeadas na área de trabalho.

## Funcionalidades

- **Mapeamento de unidades de rede:** permite mapear caminhos de rede utilizando uma letra de unidade disponível.
- **Validação de unidades:** identifica as letras já utilizadas e impede a seleção de unidades ocupadas.
- **Validação de caminhos:** verifica se o caminho de rede informado existe antes de realizar o mapeamento.
- **Identificação de unidades:** exibe as unidades de rede atualmente utilizadas e seus respectivos caminhos.
- **Criação de atalhos:** permite criar automaticamente um atalho da unidade mapeada na área de trabalho.
- **Desmapeamento de unidades:** permite remover uma unidade de rede específica ou todas as unidades de rede disponíveis na listagem.
- **Remoção de atalhos:** remove automaticamente os atalhos associados às unidades desmapeadas.
- **Tratamento de erros:** apresenta informações sobre erros encontrados durante a execução e valida as entradas fornecidas pelo usuário.

## Requisitos

- Sistema operacional Windows.
- Windows PowerShell.
- Acesso ao caminho de rede que será mapeado.
- Permissão para conectar e desconectar unidades de rede.
- Windows Script Host habilitado para criação e manipulação de atalhos.

## Como utilizar

### 1. Executar o script

Abra o PowerShell, navegue até o diretório onde o script está localizado e execute:

```powershell
.\GerenciadorUnidades.ps1
```

### 2. Menu principal

Ao iniciar, o script apresenta três opções:

```text
Digite 'mapear' para mapear um caminho de rede.
Digite 'desmapear' para desmapear um caminho de rede.
Digite 'exit' para sair.
```

### 3. Mapear uma unidade de rede

Digite `mapear` no menu principal.

O script executará as seguintes etapas:

1. Lista as unidades de rede e unidades locais atualmente utilizadas.
2. Identifica as letras disponíveis para utilização.
3. Solicita a letra que será utilizada para o mapeamento.
4. Solicita o caminho de rede (UNC).
5. Valida a existência e acessibilidade do caminho informado.
6. Solicita uma descrição para identificar a unidade.
7. Realiza o mapeamento da unidade.
8. Pergunta se deseja criar um atalho na área de trabalho.

Exemplo de caminho de rede:

```text
\\servidor\compartilhamento
```

Exemplo de execução:

```text
Unidades sendo utilizadas:
C
D
Z --- \\servidor\compartilhamento

Unidades disponíveis:
A
B
E
F
G
...

Digite a letra de uma unidade disponível: X

Digite o caminho de rede: \\servidor\projetos

Digite uma descrição: Projetos

O caminho '\\servidor\projetos' foi mapeado na unidade 'X'.

Deseja criar um atalho na area de trabalho? [S] Sim [N] Nao
```

Caso a opção `S` seja selecionada, será criado um atalho na área de trabalho com o nome informado na descrição.

### 4. Desmapear uma unidade de rede

Digite `desmapear` no menu principal.

O script apresentará as unidades atualmente utilizadas e solicitará a letra da unidade que deseja remover.

Também será disponibilizada a opção `*`, que permite selecionar todas as unidades listadas.

Exemplo:

```text
Unidades sendo utilizadas:
C
D
X --- \\servidor\projetos
Z --- \\servidor\compartilhamento

* --- Selecionar Todos

Digite a letra da unidade utilizada que deseja desmapear: X

A unidade 'X' foi selecionada.

Deseja remover a unidade 'X'? [S] Sim [N] Nao
```

Após a confirmação, a unidade será desconectada e o script verificará a existência de atalhos associados para removê-los.

**Observação:** ao selecionar `*`, todas as unidades de rede identificadas serão desmapeadas sem solicitar confirmação individual.

### 5. Encerrar a execução

Digite `exit` no menu principal para finalizar o script.

---

## Estrutura e responsabilidades das funções

O script foi desenvolvido utilizando funções independentes, separando as responsabilidades de validação, mapeamento, desmapeamento e gerenciamento de atalhos.

| Função | Responsabilidade |
|---|---|
| `Get-DriveLetterMap` | Identifica e valida letras disponíveis para mapeamento. |
| `Get-PSPath` | Solicita e valida o caminho de rede informado. |
| `Set-PSDriveValues` | Realiza o mapeamento da unidade de rede. |
| `New-Shortcut` | Cria um atalho da unidade na área de trabalho. |
| `Get-DriveLetterDesmap` | Identifica e valida a unidade que será desmapeada. |
| `Remove-PSDriveValues` | Remove uma ou todas as unidades de rede selecionadas. |
| `Remove-Shortcut` | Localiza e remove atalhos associados às unidades desmapeadas. |

### Detalhamento técnico

**Get-DriveLetterMap**

Obtém as unidades do provedor `FileSystem` por meio do `Get-PSDrive` e compara as letras utilizadas com uma lista de letras disponíveis de A até Z.

A função apresenta as unidades identificadas e solicita uma letra válida e ainda não utilizada.

**Get-PSPath**

Recebe o caminho de rede informado pelo usuário e utiliza `Test-Path` para verificar sua existência e acessibilidade.

Caso o caminho não seja encontrado, apresenta os detalhes do erro e solicita uma nova entrada.

**Set-PSDriveValues**

Utiliza o `New-PSDrive` para realizar o mapeamento da unidade.

Parâmetros utilizados:

- `-Name`: letra da unidade.
- `-Description`: descrição da unidade.
- `-PSProvider FileSystem`: define o provedor de arquivos.
- `-Root`: caminho de rede que será associado à unidade.
- `-Persist`: mantém o mapeamento no Windows.
- `-Scope Global`: disponibiliza a unidade no escopo global da sessão.

Após a execução, verifica se a unidade aparece entre as unidades do provedor `FileSystem`.

**New-Shortcut**

Utiliza o objeto COM `WScript.Shell` para criar um atalho do tipo `.lnk` na área de trabalho do usuário.

O atalho é configurado com o caminho da unidade mapeada como destino e diretório de trabalho.

**Get-DriveLetterDesmap**

Identifica as unidades atualmente utilizadas e permite selecionar uma unidade específica ou utilizar `*` para selecionar todas.

Valida a entrada do usuário antes de retornar a seleção.

**Remove-PSDriveValues**

Utiliza o comando `net use` para desconectar a unidade de rede selecionada.

Quando uma letra específica é informada, solicita confirmação antes da remoção. Quando `*` é selecionado, executa a remoção em todas as unidades de rede identificadas.

**Remove-Shortcut**

Percorre os atalhos `.lnk` existentes na área de trabalho e utiliza `WScript.Shell` para identificar seus respectivos destinos.

Quando encontra um atalho associado à unidade selecionada, realiza sua exclusão.

## Tecnologias e recursos utilizados

- **PowerShell:** linguagem utilizada para automação e gerenciamento das unidades.
- **Get-PSDrive:** identificação das unidades disponíveis no ambiente PowerShell.
- **New-PSDrive:** criação e mapeamento de unidades.
- **Test-Path:** validação de caminhos de rede.
- **Get-Item:** obtenção de informações sobre caminhos e tratamento de erros.
- **net use:** gerenciamento e desconexão de unidades de rede no Windows.
- **WScript.Shell:** automação da criação, leitura e remoção de atalhos do Windows.

## Considerações

- O mapeamento utiliza o parâmetro `-Persist`, permitindo que a unidade seja registrada no Windows.
- O caminho de rede precisa estar acessível pelo usuário que executa o script.
- A criação e remoção de atalhos são realizadas no diretório Desktop do usuário atual.
- A opção de desmapeamento geral (`*`) atua sobre todas as unidades identificadas pelo script, portanto deve ser utilizada com atenção.
- A execução depende dos recursos disponíveis no ambiente Windows e das permissões do usuário.

---

## Objetivo

O objetivo do projeto é simplificar e padronizar o gerenciamento de unidades de rede, reduzindo a necessidade de configurações manuais e oferecendo uma ferramenta interativa para automatizar operações recorrentes no ambiente Windows.
