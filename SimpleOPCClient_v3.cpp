/*
 * Sistemas Distribuidos para Automacao
 * Engenharia de Controle e Automacao
 * UFMG
 *
 * Trabalho Pratico sobre OPC e Sockets - 2024/2
 * ---------------------------------------------
 * Prof. Luiz T. S. Mendes
 *
 * Programa de simulacao do Computador de Processo
 *
 * Para a correta compilacao deste programa no WINDOWS, nao se esqueca de incluir
 * a biblioteca Winsock2 (Ws2_32.lib) no projeto ! (No Visual Studio 2010:
 * Projects->Properties->Configuration Properties->Linker->Input->A dditional Dependencies).
 *
 */

#include <atlbase.h>  
#include <iostream>
#include <ObjIdl.h>
#include <WinSock2.h>
#include <stdio.h>
#include <conio.h>

#include "opcda.h"
#include "opcerror.h"
#include "SimpleOPCClient_v3.h"
#include "SOCAdviseSink.h"
#include "SOCDataCallback.h"
#include "SOCWrapperFunctions.h"

// Para evitar "Warnings"
//#define _CRT_SECURE_NO_WARNINGS 
#define _WINSOCK_DEPRECATED_NO_WARNINGS 
#pragma warning(disable:6031)
#pragma warning(disable:6385)
#pragma comment(lib,"WS2_32")

using namespace std;

#define OPC_SERVER_NAME L"Matrikon.OPC.Simulation.1"
#define VT VT_R4

#define s       0x73
#define S       0x53
#define	ESC     0x1B
HANDLE  tecla_s;
HANDLE  tecla_ESC;
HANDLE  eventoSempreOn;

int nTecla;

#define TAMMSGDADOS   37  // 5+3+6+6+3+3+5 caracteres + 6 separadores
#define TAMMSGACK     9  // 5+3 caracteres + 1 separador
#define TAMMSGREQ     9  // 5+3 caracteres + 1 separador    
#define TAMMSGSP2     37  // 5+3+3+3+6+6+5 caracteres + 6 separadores 
#define TAMMSGACKCP   9// 5+3 caracteres + 1 separador 

#define WHITE   FOREGROUND_RED   | FOREGROUND_GREEN      | FOREGROUND_BLUE  | FOREGROUND_INTENSITY
#define HLGREEN FOREGROUND_GREEN | FOREGROUND_INTENSITY
#define HLRED   FOREGROUND_RED   | FOREGROUND_INTENSITY
#define HLBLUE  FOREGROUND_BLUE  | FOREGROUND_INTENSITY
#define YELLOW  FOREGROUND_RED   | FOREGROUND_GREEN
#define CYAN    FOREGROUND_BLUE  | FOREGROUND_GREEN      | FOREGROUND_INTENSITY
#define PURPLE  FOREGROUND_RED   | FOREGROUND_BLUE

// Função para checar erro de comunicação no socket
int CheckSocketError(int status, HANDLE hOut) {
	int erro;

	if (status == SOCKET_ERROR) {
		SetConsoleTextAttribute(hOut, HLRED);
		erro = WSAGetLastError();
		if (erro == WSAEWOULDBLOCK) {
			printf("Timeout na operacao de RECV! errno = %d - reiniciando...\n\n", erro);
			SetConsoleTextAttribute(hOut, WHITE);
			return(-1); // acarreta reinício da espera de mensagens no programa principal
		}
		else if (erro == WSAECONNABORTED) {
			printf("Conexao abortada pelo cliente TCP\n\n");
			SetConsoleTextAttribute(hOut, WHITE);
			return(-1); // acarreta reinício da espera de mensagens no programa principal
		}
		else if (erro == WSAETIMEDOUT) {
			printf("Conexao interrompida - reiniciando...\n\n");
			SetConsoleTextAttribute(hOut, WHITE);
			return(-1); // acarreta reinício da espera de mensagens no programa principal
		}
		else {
			printf("Erro de conexao! valor = %d\n\n", erro);
			return (-2); // acarreta encerramento do programa principal
		}
	}
	else if (status == 0) {
		printf("Conexao com o servidor encerrada prematuramente! status = %d\n\n", status);
		SetConsoleTextAttribute(hOut, WHITE);
		return(-1); // acarreta reinício da espera de mensagens no programa principal
	}
	else return(0);
}

// Função para encerrar conexão de sockets

void CloseConnection(SOCKET connfd) {
	closesocket(connfd);
	WSACleanup();
}

// Função de escrita no servidor
void WriteItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
{
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**)&pIOPCSyncIO);

	HRESULT* pErrors = NULL;

	HRESULT hr = pIOPCSyncIO->Write(1, &hServerItem, &varValue, &pErrors);
	if (hr != S_OK) {
		printf("Failed to send message %x.\n", hr);
		exit(0);
	}

	CoTaskMemFree(pErrors);
	pErrors = NULL;

	pIOPCSyncIO->Release();

}

// Threads
DWORD WINAPI Cliente_socketp(LPVOID);
DWORD WINAPI Cliente_socketap(LPVOID);
DWORD WINAPI Cliente_OPC(LPVOID);
DWORD WINAPI Cliente_socketHelper(LPVOID);
HANDLE mutex_socket;
HANDLE mutex_escrita;

// Variáveis socket
WSADATA     wsaData;
SOCKET      cliente, connfd;
SOCKET      servidor;
SOCKADDR_IN ServerAddr;
SOCKADDR    ClienteAddr;
int         Port = 3776;
int         status, acao;
int         contagem = 0;
char        ip[15];

/* Buffers tipos de mensagens */

// Dados de proceso e Ack
char msgdados[TAMMSGDADOS + 1] = "NNNNN$555$N.NNNN$N.NNNN$NNN$NNN$NN.NN";
char msgack[TAMMSGACK + 1] = "NNNNN$000";

// Set-points e ack 
char msgreq[TAMMSGREQ + 1] = "NNNNN$222";
char msgsp2[TAMMSGSP2+ 1] = "NNNNN$100$NNN$NNN$N.NNNN$N.NNNN$NN.NN";
char msgackcp[TAMMSGACKCP + 1] = "NNNNN$999";

int par1;
int par2;
double par3;
double par4;
int par5;

char buf[100];
char msgcode[4];
char buf_corte[100];

int nseql, nseqr, seqbuf;
int vez = 0;

// Global variables

// The OPC DA Spec requires that some constants be registered in order to use
// them. The one below refers to the OPC DA 1.0 IDataObject interface.
UINT OPC_DATA_TIME = RegisterClipboardFormat (_T("OPCSTMFORMATDATATIME"));

wchar_t ITEM_ID1[] = L"Random.Real4";
wchar_t ITEM_ID2[] = L"Random.Real8";
wchar_t ITEM_ID3[] = L"Random.Int2";
wchar_t ITEM_ID4[] = L"Random.Int4";
wchar_t ITEM_ID5[] = L"Triangle Waves.Real4";
wchar_t ITEM_ID6[] = L"Bucket Brigade.Int2";
wchar_t ITEM_ID7[] = L"Bucket Brigade.Int4";
wchar_t ITEM_ID8[] = L"Bucket Brigade.Real4";
wchar_t ITEM_ID9[] = L"Bucket Brigade.Real8";
wchar_t ITEM_ID10[] = L"Bucket Brigade.UInt1";

// Read the value of an item on an OPC server. 
//

//Leitura Opc
double* maindados;

int main(int argc, char **argv)
{
	// Handle do console
	HANDLE hOut;

	// Obtém um handle para a saída da console
	hOut = GetStdHandle(STD_OUTPUT_HANDLE);
	if (hOut == INVALID_HANDLE_VALUE)
		printf("Erro ao obter handle para a saída da console\n");
	SetConsoleTextAttribute(hOut, WHITE);

	// Inicializa objetos
	tecla_s = CreateEvent(NULL, TRUE, FALSE, "Tecla_s");              // Evento para tecla s pressionada
	tecla_ESC = CreateEvent(NULL, TRUE, FALSE, "Tecla_ESC");            // Evento para tecla ESC pressionada
	eventoSempreOn = CreateEvent(NULL, TRUE, TRUE, "EventoSempreOn");   // Evento para evento sempre on    
	mutex_socket = CreateMutex(NULL, FALSE, "Mutex_socket");            // Mutex pra acesso único do socket
	mutex_escrita = CreateMutex(NULL, FALSE, "Mutex_escrita");

	SetConsoleTextAttribute(hOut, HLGREEN);
	printf("=====================================================================================================================\n");
	printf("Inicio da Aplicacao!\n\n");
	printf("Regras para o funcionamento: \n");
	printf("1. Abra o software 'MATRIKON OPC SERVER FOR SIMULATION'.\n");
	printf("2. Abra o server e o computador de processo.\n");
	printf("3. Digite o endereco IP no terminal do computador de processo.\n");
	printf("4. A tecla ESC deve ser pressionada para finalizar a aplicacao.\n");
	printf("5. Caso queira finalizar o Servidor OPC e o Servidor de Socket, pressionar ESC no terminal da aplicacao.\n");
	printf("6. Caso queira finalizar a conexao de sockets, pressionar ESC no terminal.\n");
	printf("7. Caso queira finalizar apenas o OPC, parar o serviço do Matrikon OPC Server.\n");
	printf("=====================================================================================================================\n");
	SetConsoleTextAttribute(hOut, WHITE);

	memset(&ip, 0, sizeof(ip));
	printf("Digite o IP do computador de processo: ");
	scanf("%s", &ip);

	// Cria as Threads
	HANDLE opc = CreateThread(NULL, 0, Cliente_OPC, NULL, 0, NULL);
	/* ESPERA 2 SEGUNDOS PARA INICIAR O CLIENTE SOCKET PORQUE O OPC ESTÁ INICIANDO*/
	Sleep(2000);

	HANDLE client_socketHelper = CreateThread(NULL, 0, Cliente_socketHelper, NULL, 0, NULL);

	HANDLE handle_threads[2] = { client_socketHelper, opc };

	do
	{
		nTecla = _getch();

		switch (nTecla)
		{
		case(s):    // Minúscula
			SetEvent(tecla_s);
			break;
		case(S):    // Maiúsculo
			SetEvent(tecla_s);
			break;
		}

	} while (nTecla != ESC);

	SetEvent(tecla_ESC);    // tecla ESC foi pressionada

	//Esperar fim das threads
	WaitForMultipleObjects(2, handle_threads, TRUE, INFINITE);

	// Liberar os handles
	for (int i = 0; i < 2; i++) {
		if (handle_threads[i] != 0) CloseHandle(handle_threads[i]);
	}

	if (tecla_s != 0) CloseHandle(tecla_s);
	if (tecla_ESC != 0) CloseHandle(tecla_ESC);
	if (eventoSempreOn != 0) CloseHandle(eventoSempreOn);
	if (mutex_socket != 0) CloseHandle(mutex_socket);
	if (mutex_escrita != 0) CloseHandle(mutex_escrita);

	return(0);

}

DWORD WINAPI Cliente_socketHelper(LPVOID) {
	HANDLE eventos_esc[2] = { tecla_ESC, eventoSempreOn };    // Handle das teclas ESC e sempreon
	DWORD ret;                                                // Checar o evento
	int nTipoEvento;                                          // Ver qual o tipo do evento

	// Handle do console
	HANDLE hOut;

	// Obtém um handle para a saída da console
	hOut = GetStdHandle(STD_OUTPUT_HANDLE);
	if (hOut == INVALID_HANDLE_VALUE)
		printf("Erro ao obter handle para a saída da console\n");
	SetConsoleTextAttribute(hOut, WHITE);

	while (true) {
		ret = WaitForMultipleObjects(2, eventos_esc, FALSE, INFINITE); // Checa se memória está cheia ou eventos
		nTipoEvento = ret - WAIT_OBJECT_0;

		if (nTipoEvento == 0) break;
		else {
			/* Inicializa Winsock versão 2.2 */
			status = WSAStartup(MAKEWORD(2, 2), &wsaData);
			if (status != 0) {
				SetConsoleTextAttribute(hOut, HLRED);
				printf("Falha na inicializacao do Winsock 2! Erro  = %d\n", WSAGetLastError());
				SetConsoleTextAttribute(hOut, WHITE);
				WSACleanup();
				exit(0);
			}

			/* Cria socket */
			cliente = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
			if (cliente == INVALID_SOCKET) {
				SetConsoleTextAttribute(hOut, HLRED);
				status = WSAGetLastError();
				if (status == WSAENETDOWN)
					printf("Rede ou servidor de sockets inacessíveis!\n");
				else
					printf("Falha na rede: codigo de erro = %d\n", status);
				SetConsoleTextAttribute(hOut, WHITE);
				WSACleanup();
				exit(0);
			}

			/* Permite a possibilidade de reuso deste socket, de forma que,   */
			/* se uma instância anterior deste programa tiver sido encerrada  */
			/* com CTRL-C por exemplo, não ocorrera' o erro "10048" ("address */
			/* already in use") na operacao de BIND                           */
			int optval = 1;
			setsockopt(cliente, SOL_SOCKET, SO_REUSEADDR, (const char*)&optval, sizeof(optval));

			/* Inicializa estrutura sockaddr_in */
			memset(&ServerAddr, 0, sizeof(ServerAddr));
			ServerAddr.sin_family = AF_INET;
			ServerAddr.sin_addr.s_addr = inet_addr(ip);
			ServerAddr.sin_port = htons(Port);

			/* Vincula o socket ao endereco e porta especificados */
			status = connect(cliente, (SOCKADDR*)&ServerAddr, sizeof(ServerAddr));
			if (status == SOCKET_ERROR) {
				SetConsoleTextAttribute(hOut, HLRED);
				if (WSAGetLastError() == 10049 || WSAGetLastError() == 10065 || WSAGetLastError() == 10051) {
					printf("IP invalido digitado para a conexao\n");
				}
				else printf("Falha na funcao connect! Erro  = %d\n", WSAGetLastError());

				SetConsoleTextAttribute(hOut, WHITE);
				WSACleanup();
				exit(0);
			}
			else {
				WaitForSingleObject(mutex_escrita, 3000);
				printf("Conexao com o servidor efetuada!\n\n");
				ReleaseMutex(mutex_escrita);
			}


			/* DISPARA AS THREADS PERIÓDICA E APERIÓDICA E ESPERA ERRO DE CONEXÃO */
			HANDLE client_socketp = CreateThread(NULL, 0, Cliente_socketp, NULL, 0, NULL);
			HANDLE client_socketap = CreateThread(NULL, 0, Cliente_socketap, NULL, 0, NULL);
			HANDLE eventos_erro[3] = { tecla_ESC, client_socketp, client_socketap };    // Handle das threads e ESC

			/* ESPERA */
			ret = WaitForMultipleObjects(3, eventos_erro, FALSE, INFINITE); // Checa eventos
			nTipoEvento = ret - WAIT_OBJECT_0;

			if (nTipoEvento == 0) {
				/* FINALIZA THREADS PARA BREAK */
				CloseHandle(client_socketp);
				CloseHandle(client_socketap);
				break;
			}

			else {
				/* FINALIZA THREADS E TENTA RECONECTAR */
				TerminateThread(client_socketap, 0);
				TerminateThread(client_socketp, 0);
				CloseHandle(client_socketp);
				CloseHandle(client_socketap);
				CloseConnection(cliente);
				contagem = 0; // Reseta a contagem
				SetConsoleTextAttribute(hOut, HLRED);
				printf("Esperando 10 segundos para reconectar...\n\n");
				SetConsoleTextAttribute(hOut, WHITE);
				Sleep(10000);
			}
		}
	}

	WaitForSingleObject(mutex_escrita, 3000);
	SetConsoleTextAttribute(hOut, HLRED);
	printf("Conexao com servidor de processo encerrado...\n");
	SetConsoleTextAttribute(hOut, WHITE);
	ReleaseMutex(mutex_escrita);

	CloseConnection(cliente);
	return 0;
}

DWORD WINAPI Cliente_socketp(LPVOID) {
	// Eventosss
	HANDLE eventos_esc[2] = { tecla_ESC, eventoSempreOn };    // Handle das teclas ESC e sempreon
	DWORD ret;                                                // Checar o evento
	int nTipoEvento;                                          // Ver qual o tipo do evento

	// Handle do console
	HANDLE hOut;

	// Obtém um handle para a saída da console
	hOut = GetStdHandle(STD_OUTPUT_HANDLE);
	if (hOut == INVALID_HANDLE_VALUE)
		printf("Erro ao obter handle para a saída da console\n");
	SetConsoleTextAttribute(hOut, WHITE);

	// *** Neste ponto o cliente envia requisições ao servidor e recebe as respostas ***.
	while (true) {
		ret = WaitForMultipleObjects(2, eventos_esc, FALSE, INFINITE); // Checa se memória está cheia ou eventos
		nTipoEvento = ret - WAIT_OBJECT_0;

		if (nTipoEvento == 0) break;
		else {
			/* ESPERA MUTEX */
			WaitForSingleObject(mutex_socket, INFINITE);

			/* CHECAR CONTAGEM */
			if (++contagem > 99999) contagem = 1;

			/* Envia mensagem de dados */
			memset(buf, 0, sizeof(buf));

			sprintf(buf, "%05d$555$%06f$%06f$%03d$%03d$%05f", contagem, maindados[0], maindados[1], maindados[2], maindados[3], maindados[4]);

			memcpy(msgdados, buf, sizeof(msgdados));

			status = send(cliente, msgdados, TAMMSGDADOS, 0);
			if ((acao = CheckSocketError(status, hOut)) != 0) break;

			WaitForSingleObject(mutex_escrita, 3000);
			SetConsoleTextAttribute(hOut, HLGREEN);
			printf("(E) Mensagem de dados enviada ao servidor:\n%s\n", msgdados);
			SetConsoleTextAttribute(hOut, WHITE);
			ReleaseMutex(mutex_escrita);

			/* Recebe o ACK */
			memset(buf, 0, sizeof(buf));
			status = recv(cliente, buf, TAMMSGACK, 0);
			if ((acao = CheckSocketError(status, hOut)) != 0) break;

			/* CHECAR CONTAGEM */
			if (++contagem > 99999) contagem = 1;

			memcpy(msgack, buf, TAMMSGACK + 1);

			WaitForSingleObject(mutex_escrita, 3000);
			SetConsoleTextAttribute(hOut, YELLOW);
			printf("(R) Mensagem de ACK recebida do servidor:\n%s\n\n", msgack);
			SetConsoleTextAttribute(hOut, WHITE);
			ReleaseMutex(mutex_escrita);

			/* LIBERA MUTEX */
			ReleaseMutex(mutex_socket);

			/* RESETAR TECLA S */
			ResetEvent(tecla_s);

			/* ESPERA 2 SEGUNDOS */
			Sleep(2000);
		}
	}
	return 0;
}

DWORD WINAPI Cliente_socketap(LPVOID) {
	// Eventosss
	HANDLE eventos_s[2] = { tecla_ESC, tecla_s };         // Handle das teclas ESC e 
	DWORD ret;  // Checar o evento
	int nTipoEvento;    // Ver qual o tipo do evento

	// Handle do console
	HANDLE hOut;

	// Obtém um handle para a saída da console
	hOut = GetStdHandle(STD_OUTPUT_HANDLE);
	if (hOut == INVALID_HANDLE_VALUE)
		printf("Erro ao obter handle para a saída da console\n");
	SetConsoleTextAttribute(hOut, WHITE);

	char* token;

	// *** Neste ponto o cliente envia requisições ao servidor e recebe as respostas ***.
	while (true) {
		ret = WaitForMultipleObjects(2, eventos_s, FALSE, INFINITE); // Checa se memória está cheia ou eventos
		nTipoEvento = ret - WAIT_OBJECT_0;

		if (nTipoEvento == 0) break;
		else {
			/* ESPERA MUTEX */
			WaitForSingleObject(mutex_socket, INFINITE);

			/* CHECAR CONTAGEM */
			if (++contagem > 99999) contagem = 1;

			/* Envia mensagem de requisição */
			memset(buf, 0, sizeof(buf));
			sprintf(buf, "%05d", contagem);
			memcpy(msgreq, buf, 5);
			status = send(cliente, msgreq, TAMMSGREQ, 0);
			if ((acao = CheckSocketError(status, hOut)) != 0) break;

			WaitForSingleObject(mutex_escrita, 3000);
			SetConsoleTextAttribute(hOut, HLBLUE);
			printf("(E) Mensagem de requisicao enviada ao servidor:\n%s\n", msgreq);
			SetConsoleTextAttribute(hOut, WHITE);
			ReleaseMutex(mutex_escrita);

			/* Recebe a resposta */
			memset(buf, 0, sizeof(buf));
			status = recv(cliente, buf, TAMMSGSP2, 0);
			if ((acao = CheckSocketError(status, hOut)) != 0) break;

			/* CHECAR CONTAGEM */
			if (++contagem > 99999) contagem = 1;

			memcpy(msgsp2, buf, TAMMSGSP2 + 1);

			WaitForSingleObject(mutex_escrita, 3000);
			SetConsoleTextAttribute(hOut, CYAN);
			printf("(R) Mensagem de set-point recebido do servidor:\n%s\n", msgsp2);
			SetConsoleTextAttribute(hOut, WHITE);
			ReleaseMutex(mutex_escrita);

			memset(buf, 0, sizeof(buf));
			strcpy(buf, msgsp2);

			// Coloca na memória cada parâmetro
			token = strtok(buf, "$");        // ID
			token = strtok(NULL, "$");       // codigo

			token = strtok(NULL, "$");       // p1
			par1 = atof(token);

			token = strtok(NULL, "$");       // p2
			par2 = atof(token);

			token = strtok(NULL, "$");       // p3
			par3 = atof(token);

			token = strtok(NULL, "$");       // p4
			par4 = atof(token);

			token = strtok(NULL, "\0");      // p5
			par5 = atoi(token);

			/* CHECAR CONTAGEM */
			if (++contagem > 99999) contagem = 1;

			/* Devolve ACK */
			memset(buf, 0, sizeof(buf));
			sprintf(buf, "%05d", contagem);
			memcpy(msgackcp, buf, 5);
			status = send(cliente, msgackcp, TAMMSGACKCP, 0);
			if ((acao = CheckSocketError(status, hOut)) != 0) break;

			WaitForSingleObject(mutex_escrita, 3000);
			SetConsoleTextAttribute(hOut, YELLOW);
			printf("(E) Mensagem de ACK enviada ao servidor:\n%s\n\n", msgackcp);
			SetConsoleTextAttribute(hOut, WHITE);
			ReleaseMutex(mutex_escrita);

			/* LIBERA MUTEX */
			ReleaseMutex(mutex_socket);

			/* RESETAR TECLA S */
			ResetEvent(tecla_s);
		}
	}

	return 0;
}

DWORD WINAPI Cliente_OPC(LPVOID)
{

	// Handle do console
	HANDLE hOut;

	// Obtém um handle para a saída da console
	hOut = GetStdHandle(STD_OUTPUT_HANDLE);
	if (hOut == INVALID_HANDLE_VALUE) {
		SetConsoleTextAttribute(hOut, HLRED);
		printf("Erro ao obter handle para a saída da console\n");
		SetConsoleTextAttribute(hOut, WHITE);
	}

	// Eventos
	HANDLE eventos[2] = { tecla_ESC, eventoSempreOn };         // Handle das teclas ESC sempre on
	DWORD ret;  // Checar o evento
	int nTipoEvento;    // Ver qual o tipo do evento

	IOPCServer* pIOPCServer = NULL;   //pointer to IOPServer interface
	IOPCItemMgt* pIOPCItemMgt = NULL; //pointer to IOPCItemMgt interface

	OPCHANDLE hServerGroup;  // server handle to the group

	OPCHANDLE hServerItem1;  // server handle to the item
	OPCHANDLE hServerItem2;  // server handle to the item
	OPCHANDLE hServerItem3;  // server handle to the item
	OPCHANDLE hServerItem4;  // server handle to the item
	OPCHANDLE hServerItem5;  // server handle to the item
	OPCHANDLE hServerItem6;  // server handle to the item
	OPCHANDLE hServerItem7;  // server handle to the item
	OPCHANDLE hServerItem8;  // server handle to the item
	OPCHANDLE hServerItem9;  // server handle to the item
	OPCHANDLE hServerItem10;  // server handle to the item

	char buf[100];

	// Have to be done before using microsoft COM library:
	//printf("Initializing the COM environment...\n");
	CoInitialize(NULL);

	// Let's instantiante the IOPCServer interface and get a pointer of it:
	//printf("Instantiating the MATRIKON OPC Server for Simulation...\n");
	pIOPCServer = InstantiateServer(OPC_SERVER_NAME);

	// Add the OPC group the OPC server and get an handle to the IOPCItemMgt
	//interface:
	//printf("Adding a group in the INACTIVE state for the moment...\n");
	AddTheGroup(pIOPCServer, pIOPCItemMgt, hServerGroup);

	// Add the OPC item. First we have to convert from wchar_t* to char*
	// in order to print the item name in the console.
	size_t m;
	wcstombs_s(&m, buf, 100, ITEM_ID1, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID2, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID3, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID4, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID5, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID6, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID7, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID8, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID9, _TRUNCATE);
	wcstombs_s(&m, buf, 100, ITEM_ID10, _TRUNCATE);


	AddTheItem(pIOPCItemMgt, hServerItem1, hServerItem2, hServerItem3, hServerItem4, hServerItem5, hServerItem6, hServerItem7, hServerItem8, hServerItem9, hServerItem10);

	// Lendo de forma assíncrona

	int bRet;
	MSG msg;

	//Callback
	IConnectionPoint* pIConnectionPoint = NULL; //pointer to IConnectionPoint Interface
	DWORD dwCookie = 0;
	SOCDataCallback* pSOCDataCallback = new SOCDataCallback();
	pSOCDataCallback->AddRef();

	SetDataCallback(pIOPCItemMgt, pSOCDataCallback, pIConnectionPoint, &dwCookie);

	SetGroupActive(pIOPCItemMgt);

	// Lendo os dados, gerando um fluxo de mensagens

	VARIANT aux;
	VariantInit(&aux);

	while (true) {

		ret = WaitForMultipleObjects(2, eventos, FALSE, INFINITE); // Checa se memória está cheia ou eventos
		nTipoEvento = ret - WAIT_OBJECT_0;

		if (nTipoEvento == 0) break;
		else {

			bRet = GetMessage(&msg, NULL, 0, 0);
			if (!bRet) {
				printf("Failed to get windows message! Error code = %d\n", GetLastError());
				exit(0);
			}
			TranslateMessage(&msg); // This call is not really needed ...
			DispatchMessage(&msg);  // ... but this one is!

			maindados = pSOCDataCallback->leitura_dados();

			// Leitura do Servidor OPC
			WaitForSingleObject(mutex_escrita, 3000);
			SetConsoleTextAttribute(hOut, PURPLE);
			printf("\n=======  Leitura do Servidor OPC  =======\n");
			printf("[OPC] Concentracao de HCI no tanque 1:      %f\n", maindados[0]);
			printf("[OPC] Concentracao de HCI no tanque 2:       %f\n", maindados[1]);
			printf("[OPC] Temperatura no tanque 1:   %f\n", maindados[2]);
			printf("[OPC] Temperatura no tanque 2:   %f\n", maindados[3]);
			printf("[OPC] Velocidade da tira:   %f\n\n", maindados[4]);
			SetConsoleTextAttribute(hOut, WHITE);
			ReleaseMutex(mutex_escrita);


			if (strcmp(msgsp2, "") != 0) {


				// Escrita no Servidor OPC
				aux.vt = VT_I2;
				aux.intVal = par1;
				WriteItem(pIOPCItemMgt, hServerItem6, aux);

				aux.vt = VT_I4;
				aux.intVal = par2;
				WriteItem(pIOPCItemMgt, hServerItem7, aux);

				aux.vt = VT_R4;
				aux.fltVal = par3;
				WriteItem(pIOPCItemMgt, hServerItem8, aux);

				aux.vt = VT_R8;
				aux.dblVal = par4;
				WriteItem(pIOPCItemMgt, hServerItem9, aux);

				aux.vt = VT_UI1;
				aux.uiVal = par5;
				WriteItem(pIOPCItemMgt, hServerItem10, aux);

				WaitForSingleObject(mutex_escrita, 3000);
				SetConsoleTextAttribute(hOut, PURPLE);
				printf("=======  Escrita no Servidor OPC  =======\n");
				printf("[OPC] Set-point de temperatura no tanque 1: %d\n", par1);
				printf("[OPC] Set-point de temperatura no tanque 2: %ld\n", par2);
				printf("[OPC] Set-point de concentracao acida no tanque 1: %f\n", par3);
				printf("[OPC] Set-point de concentracao acida no tanque 2: %f\n", par4);
				printf("[OPC] Set-point de velocidade da tira: %u\n\n", par5);

				SetConsoleTextAttribute(hOut, WHITE);
				ReleaseMutex(mutex_escrita);

				// Zera o valor da mensagem de parâmetros para evitar impressões repetitivas
				memset(msgsp2, 0, sizeof(msgsp2));

			}
		}
	}

	CancelDataCallback(pIConnectionPoint, dwCookie);
	//pIConnectionPoint->Release();
	pSOCDataCallback->Release();

	// Remove the OPC item:
	RemoveItem(pIOPCItemMgt, hServerItem1);
	RemoveItem(pIOPCItemMgt, hServerItem2);
	RemoveItem(pIOPCItemMgt, hServerItem3);
	RemoveItem(pIOPCItemMgt, hServerItem4);
	RemoveItem(pIOPCItemMgt, hServerItem5);
	RemoveItem(pIOPCItemMgt, hServerItem6);
	RemoveItem(pIOPCItemMgt, hServerItem7);
	RemoveItem(pIOPCItemMgt, hServerItem8);
	RemoveItem(pIOPCItemMgt, hServerItem9);
	RemoveItem(pIOPCItemMgt, hServerItem10);
	// Remove the OPC group:
	//printf("Removing the OPC group object...\n");
	pIOPCItemMgt->Release();
	RemoveGroup(pIOPCServer, hServerGroup);

	// release the interface references:
	//printf("Removing the OPC server object...\n");
	pIOPCServer->Release();

	//close the COM library:
	WaitForSingleObject(mutex_escrita, 3000);
	SetConsoleTextAttribute(hOut, HLRED);
	printf("Ambiente OPC finalizado...\n");
	SetConsoleTextAttribute(hOut, WHITE);
	ReleaseMutex(mutex_escrita);

	CoUninitialize();

	closesocket(servidor);

	return 0;
}

////////////////////////////////////////////////////////////////////
// Instantiate the IOPCServer interface of the OPCServer
// having the name ServerName. Return a pointer to this interface
//
IOPCServer* InstantiateServer(wchar_t ServerName[])
{
	CLSID CLSID_OPCServer;
	HRESULT hr;

	// get the CLSID from the OPC Server Name:
	hr = CLSIDFromString(ServerName, &CLSID_OPCServer);
	_ASSERT(!FAILED(hr));


	//queue of the class instances to create
	LONG cmq = 1; // nbr of class instance to create.
	MULTI_QI queue[1] =
	{ {&IID_IOPCServer,
	NULL,
	0} };

	//Server info:
	//COSERVERINFO CoServerInfo =
	//{
	//	/*dwReserved1*/ 0,
	//	/*pwszName*/ REMOTE_SERVER_NAME,
	//	/*COAUTHINFO*/  NULL,
	//	/*dwReserved2*/ 0
	//}; 

	// create an instance of the IOPCServer
	hr = CoCreateInstanceEx(CLSID_OPCServer, NULL, CLSCTX_SERVER,
		/*&CoServerInfo*/NULL, cmq, queue);
	_ASSERT(!hr);

	// return a pointer to the IOPCServer interface:
	return(IOPCServer*)queue[0].pItf;
}


/////////////////////////////////////////////////////////////////////
// Add group "Group1" to the Server whose IOPCServer interface
// is pointed by pIOPCServer. 
// Returns a pointer to the IOPCItemMgt interface of the added group
// and a server opc handle to the added group.
//
void AddTheGroup(IOPCServer* pIOPCServer, IOPCItemMgt*& pIOPCItemMgt,
	OPCHANDLE& hServerGroup)
{
	DWORD dwUpdateRate = 0;
	OPCHANDLE hClientGroup = 0;

	// Add an OPC group and get a pointer to the IUnknown I/F:
	HRESULT hr = pIOPCServer->AddGroup(/*szName*/ L"Group1",
		/*bActive*/ FALSE,
		/*dwRequestedUpdateRate*/ 1000,
		/*hClientGroup*/ hClientGroup,
		/*pTimeBias*/ 0,
		/*pPercentDeadband*/ 0,
		/*dwLCID*/0,
		/*phServerGroup*/&hServerGroup,
		&dwUpdateRate,
		/*riid*/ IID_IOPCItemMgt,
		/*ppUnk*/ (IUnknown**)&pIOPCItemMgt);
	_ASSERT(!FAILED(hr));
}



//////////////////////////////////////////////////////////////////
// Add the Item ITEM_ID to the group whose IOPCItemMgt interface
// is pointed by pIOPCItemMgt pointer. Return a server opc handle
// to the item.

void AddTheItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE& hServerItem1, OPCHANDLE& hServerItem2, OPCHANDLE& hServerItem3, OPCHANDLE& hServerItem4, OPCHANDLE& hServerItem5, OPCHANDLE& hServerItem6, OPCHANDLE& hServerItem7, OPCHANDLE& hServerItem8, OPCHANDLE& hServerItem9, OPCHANDLE& hServerItem10)
{
	HRESULT hr;

	// Array of items to add:
	OPCITEMDEF ItemArray[10] =
	{ {
			/*szAccessPath*/ L"",
			/*szItemID*/ ITEM_ID1,
			/*bActive*/ TRUE,
			/*hClient*/ 1,
			/*dwBlobSize*/ 0,
			/*pBlob*/ NULL,
			/*vtRequestedDataType*/ VT,
			/*wReserved*/0
			},
			{
				/*szAccessPath*/ L"",
				/*szItemID*/ ITEM_ID2,
				/*bActive*/ TRUE,
				/*hClient*/ 2,
				/*dwBlobSize*/ 0,
				/*pBlob*/ NULL,
				/*vtRequestedDataType*/ VT,
				/*wReserved*/0
				},
				{
					/*szAccessPath*/ L"",
					/*szItemID*/ ITEM_ID3,
					/*bActive*/ TRUE,
					/*hClient*/ 3,
					/*dwBlobSize*/ 0,
					/*pBlob*/ NULL,
					/*vtRequestedDataType*/ VT,
					/*wReserved*/0
					},
					{
						/*szAccessPath*/ L"",
						/*szItemID*/ ITEM_ID4,
						/*bActive*/ TRUE,
						/*hClient*/ 4,
						/*dwBlobSize*/ 0,
						/*pBlob*/ NULL,
						/*vtRequestedDataType*/ VT,
						/*wReserved*/0
						},
						{
							/*szAccessPath*/ L"",
							/*szItemID*/ ITEM_ID5,
							/*bActive*/ TRUE,
							/*hClient*/ 5,
							/*dwBlobSize*/ 0,
							/*pBlob*/ NULL,
							/*vtRequestedDataType*/ VT,
							/*wReserved*/0
							},
							{
								/*szAccessPath*/ L"",
								/*szItemID*/ ITEM_ID6,
								/*bActive*/ TRUE,
								/*hClient*/ 6,
								/*dwBlobSize*/ 0,
								/*pBlob*/ NULL,
								/*vtRequestedDataType*/ VT,
								/*wReserved*/0
								},
								{
									/*szAccessPath*/ L"",
									/*szItemID*/ ITEM_ID7,
									/*bActive*/ TRUE,
									/*hClient*/ 7,
									/*dwBlobSize*/ 0,
									/*pBlob*/ NULL,
									/*vtRequestedDataType*/ VT,
									/*wReserved*/0
									},
										{
											/*szAccessPath*/ L"",
											/*szItemID*/ ITEM_ID8,
											/*bActive*/ TRUE,
											/*hClient*/ 8,
											/*dwBlobSize*/ 0,
											/*pBlob*/ NULL,
											/*vtRequestedDataType*/ VT,
											/*wReserved*/0
											},
											{
												/*szAccessPath*/ L"",
												/*szItemID*/ ITEM_ID9,
												/*bActive*/ TRUE,
												/*hClient*/ 9,
												/*dwBlobSize*/ 0,
												/*pBlob*/ NULL,
												/*vtRequestedDataType*/ VT,
												/*wReserved*/0
												},
												{
													/*szAccessPath*/ L"",
													/*szItemID*/ ITEM_ID10,
													/*bActive*/ TRUE,
													/*hClient*/ 10,
													/*dwBlobSize*/ 0,
													/*pBlob*/ NULL,
													/*vtRequestedDataType*/ VT,
													/*wReserved*/0
													},
	};

	//Add Result:
	OPCITEMRESULT* pAddResult = NULL;
	HRESULT* pErrors = NULL;

	// Add an Item to the previous Group:
	hr = pIOPCItemMgt->AddItems(10, ItemArray, &pAddResult, &pErrors);
	if (hr != S_OK) {
		printf("Failed call to AddItems function. Error code = %x\n", hr);
		exit(0);
	}

	// Server handle for the added item:
	hServerItem1 = pAddResult[0].hServer;
	hServerItem2 = pAddResult[1].hServer;
	hServerItem3 = pAddResult[2].hServer;
	hServerItem4 = pAddResult[3].hServer;
	hServerItem5 = pAddResult[4].hServer;
	hServerItem6 = pAddResult[5].hServer;
	hServerItem7 = pAddResult[6].hServer;
	hServerItem8 = pAddResult[7].hServer;
	hServerItem9 = pAddResult[8].hServer;
	hServerItem10 = pAddResult[9].hServer;

	// release memory allocated by the server:
	CoTaskMemFree(pAddResult->pBlob);

	CoTaskMemFree(pAddResult);
	pAddResult = NULL;

	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

///////////////////////////////////////////////////////////////////////////////
// Read from device the value of the item having the "hServerItem" server 
// handle and belonging to the group whose one interface is pointed by
// pGroupIUnknown. The value is put in varValue. 
//
void ReadItem(IUnknown* pGroupIUnknown, OPCHANDLE hServerItem, VARIANT& varValue)
{
	// value of the item:
	OPCITEMSTATE* pValue = NULL;

	//get a pointer to the IOPCSyncIOInterface:
	IOPCSyncIO* pIOPCSyncIO;
	pGroupIUnknown->QueryInterface(__uuidof(pIOPCSyncIO), (void**)&pIOPCSyncIO);

	// read the item value from the device:
	HRESULT* pErrors = NULL; //to store error code(s)
	HRESULT hr = pIOPCSyncIO->Read(OPC_DS_DEVICE, 1, &hServerItem, &pValue, &pErrors);
	_ASSERT(!hr);
	_ASSERT(pValue != NULL);

	varValue = pValue[0].vDataValue;

	//Release memeory allocated by the OPC server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;

	CoTaskMemFree(pValue);
	pValue = NULL;

	// release the reference to the IOPCSyncIO interface:
	pIOPCSyncIO->Release();
}

///////////////////////////////////////////////////////////////////////////
// Remove the item whose server handle is hServerItem from the group
// whose IOPCItemMgt interface is pointed by pIOPCItemMgt
//
void RemoveItem(IOPCItemMgt* pIOPCItemMgt, OPCHANDLE hServerItem)
{
	// server handle of items to remove:
	OPCHANDLE hServerArray[1];
	hServerArray[0] = hServerItem;

	//Remove the item:
	HRESULT* pErrors; // to store error code(s)
	HRESULT hr = pIOPCItemMgt->RemoveItems(1, hServerArray, &pErrors);
	_ASSERT(!hr);

	//release memory allocated by the server:
	CoTaskMemFree(pErrors);
	pErrors = NULL;
}

////////////////////////////////////////////////////////////////////////
// Remove the Group whose server handle is hServerGroup from the server
// whose IOPCServer interface is pointed by pIOPCServer
//
void RemoveGroup(IOPCServer* pIOPCServer, OPCHANDLE hServerGroup)
{
	// Remove the group:
	HRESULT hr = pIOPCServer->RemoveGroup(hServerGroup, FALSE);
	if (hr != S_OK) {
		if (hr == OPC_S_INUSE)
			printf("Failed to remove OPC group: object still has references to it.\n");
		else printf("Failed to remove OPC group. Error code = %x\n", hr);
		exit(0);
	}
}
