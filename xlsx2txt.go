package main

import (
	"C"
)
import (
	"bufio"
	"fmt"
	"os"
	"strings"

	"github.com/tealeg/xlsx"
)

func retError(err error) *C.char {
	return C.CString(fmt.Sprintf(`%v`, err))
}

//export GetInt
func GetInt() int32 {
	return 1
}

//export SaveAsTXT
func SaveAsTXT(pFileNameXLSX *C.char, pFileNameTXT *C.char, pColsName *C.char) *C.char {
	fileXLSX, err := xlsx.OpenFile(C.GoString(pFileNameXLSX))

	if err != nil {
		return retError(err)
	}

	file, err := os.OpenFile(C.GoString(pFileNameTXT), os.O_CREATE|os.O_WRONLY, 0644)

	if err != nil {
		return retError(err)
	}

	datawriter := bufio.NewWriter(file)

	vColsName := strings.ToUpper(C.GoString(pColsName))
	vColsName = strings.ReplaceAll(vColsName, " ", "")

	var vColsToPrint []bool

	vLinha := 0
	vColuna := 0
	vCanPrint := false
	vFindCol := vColsName != ""
	vTextoLinha := ""

	for _, sheet := range fileXLSX.Sheets {

		for nRow := 0; nRow < len(sheet.Rows); nRow++ {
			vColuna = 0
			vTextoLinha = ""

			if vFindCol {
				vColsToPrint = nil
			}

			for _, cell := range sheet.Rows[nRow].Cells {
				cell.NumFmt = "##########,#########0.000000000;##,##00.00"
				vCol := cell.String()

				if vFindCol {
					vLinha = 0

					if strings.Trim(vCol, " ") == "" || !strings.Contains(vColsName, fmt.Sprintf("%s;", strings.ToUpper(strings.ReplaceAll(vCol, " ", "")))) {
						vColsToPrint = append(vColsToPrint, false)
						vColuna++
						continue
					}

					vFindCol = false
				}

				if vLinha == 0 {
					vCanPrint = vColsName == "" || strings.Contains(vColsName, fmt.Sprintf("%s;", strings.ToUpper(strings.ReplaceAll(vCol, " ", ""))))
					vColsToPrint = append(vColsToPrint, vCanPrint)
				}

				if vColuna > len(vColsToPrint)-1 {
					nRow--
					vLinha = 0
					vFindCol = vColsName != ""
					vColsToPrint = nil

					break
				}

				if vColsToPrint[vColuna] {
					vTextoLinha = vTextoLinha + vCol + ";"
					//_, err = datawriter.WriteString(vCol + ";")

					/*if err != nil {
						return retError(err)
					}*/
				}

				vColuna++
			}

			if vColsToPrint == nil {
				continue
			}

			if !vFindCol {
				_, err = datawriter.WriteString(vTextoLinha + "\n")
			}

			if err != nil {
				return retError(err)
			}

			vLinha++

		}

		//Até o momento, desenvolvido considerando que o arquivo XLSX terá apenas uma planilha
		break
	}

	datawriter.Flush()
	file.Close()

	return C.CString("sucess")
}

func main() {
	const cols string = "DATADORECEBIMENTO;DATAORIGINALDAVENDA;VALORBRUTODAPARCELAORIGINAL;VALORBRUTODAPARCELAATUALIZADO;TAXAMDR;VALORMDRDESCONTADO;VALORLÍQUIDODAPARCELA;NSU/CV;NÚMERODOPEDIDO;NÚMERODAAUTORIZAÇÃO;ESTABELECIMENTO;NÚMERODOCARTÃO;MODALIDADE;BANDEIRA;NÚMERODEPARCELAS;PARCELA;BANCO;AGÊNCIA;CONTA-CORRENTE;CANCELAMENTO/CONTESTAÇÃO;DATADOCANCELAMENTO;STATUS;"
	fmt.Printf(C.GoString(SaveAsTXT(C.CString("teste2.xlsx"), C.CString("teste.txt"), C.CString(cols))))
}

//fmt.Printf(C.GoString(SaveAsTXT(C.CString("teste.xlsx"), C.CString("teste.txt"), C.CString("DATAPEDIDO;DATAPAGAMENTO;DATAESTORNO;DATALIBERAÇÃO;DATAPREVISTAPGTO;REF.PEDIDO;ENTREGA;TIPO;VALOR;"))))

//"data do recebimento;data original da venda;valor bruto da parcela original;valor bruto da parcela atualizado;taxa MDR;valor MDR descontado;valor líquido da parcela;NSU/CV;número do pedido;número da autorização;estabelecimento;número do cartão;modalidade;bandeira;número de parcelas;parcela;banco;agência;conta-corrente;cancelamento/contestação;data do cancelamento;status;"
