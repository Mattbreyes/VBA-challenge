sub stonks():

    For each ws in worksheets
    
    'formatting and definitions
        
        'variable definitions

        dim rowcount as long
        rowcount = ws.cells(rows.count, 1).end(xlup).row

        dim tickers as long
        tickers = 2

        dim currentticker as string
        currentticker = ws.cells(2,1).value

        dim volsum as longlong
        volsum = 0

        dim openprice as double
        openprice = ws.cells(2,3).value

        dim closeprice as double
        closeprice = 0

        dim greatestticker as string
        greaetesticker = 0

        dim lowestticker as string
        lowestticker = 0

        dim tickervolume as string
        tickervolume = ws.cells(2,9).value

        
        'bonus 

        dim greatestincrease as string
        greatestincrease = 0

        dim greatestdecrease as double
        greatestdecrease = 0

        dim greatestvolume as longlong
        greatestvolume = 0




        'column headers

        ws.range("I1").value = "Ticker"
        ws.range("J1").value = "Yearly Change"
        ws.range("K1").value = "Percentage Change"
        ws.range("L1").value = "Total Stock Volume"

        ws.range("P1").value = "Ticker"
        ws.range("Q1").value = "Value"
        ws.range("O2").value = "Greatest % Increase"
        ws.range("O4").value = "Greatest % Decrease"
        ws.range("O4").value = "Greatest Total Volume"
        
        ws.columns("O").autofit

    
    
    'loop start

        for i = 2 to rowcount

            'add current to previous volume
            volsum = volsum + ws.cells(i,7).value


            'if next ticker is diff then print current values
            if currentticker <> ws.cells(i + 1, 1).value then

                'ticker format
                ws.cells(tickers, 9).value = currentticker

                'yearly change format
                ws.cells(tickers, 10).numberformat = "0.00"

                'percent change format
                ws.cells(tickers, 11).value = formatpercent((closeprice / openprice) -1)

                'total stock volume format
                ws.cells(tickers, 12).value = volsum


                closeprice = ws.cells(i,6).value

                
                'yearly change calculations

                dim yearlychange as long
                yearlychange = formatnumber(closeprice - openprice, 2)
                
                ws.cells(tickers, 10).value = yearlychange

                ws.cells(tickers, 10).numberformat = "0.00"

                currentticker = ws.cells(i + 1, 1).value

                if  yearlychange > 0 then 

                    ws.cells(tickers, 10).interior.colorindex = 4

                else

                    ws.cells(tickers, 10).interior.colorindex = 3
                
                end if

                
                
                'percentage change calculations

                if ws.cells(tickers, 11).value > greatestincrease then

                    greatestincrease = ws.cells(tickers, 11).value

                    greatestticker = ws.cells(tickers, 9).value

                else

                    greatestdecrease = ws.cells(tickers, 11).value

                    lowestticker = ws.cells(tickers, 9).value

                end if

                'total stock volume calculations

                if ws.cells(tickers, 12) > greatestvolume then

                    greatestvolume = ws.cells(tickers, 12).Value
                    
                    tickervolume = ws.cells(tickers, 9).value

                end if

            openprice = ws.cells(i + 1, 3).value
            volsum = 0
            tickers = tickers + 1
            
            end if
        next i

    ws.columns("J:L").autofit


    'bonus table

    ws.range("P2").value = greatestticker
    ws.range("P3").value = lowestticker
    ws.range("P4").value = tickervolume
    
    ws.range("Q2").value = formatpercent(greatestincrease)
    ws.range("Q3").value = formatpercent(greatestdecrease)
    ws.range("Q4").value = greatestvolume

    next ws

end sub







