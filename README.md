# SandP_500-algo_trade
A more cleaner code is coming in the next update. <br>
Collect data from S &amp; P top 500 companies listed in NASDAQ and create investment strategies based on portfolio size.
Thanks to https://iexcloud.io/docs/api/ for providing with the data.
Thanks to FreeCodeCamp for provinding the IEX API CLOUD TOKEN.
Only quant_trading.py needs to be run by the user.
Three different algorithmic trading strategies are discussed: Equal Weights Strategy, Quantitative Momentum Strategy and Quantitative Value Strategy
For equal weights with individual requests, the data is collected only for 1st 100 results.
For equal weights with specific, the value currently is set to 'AAPL'. Change the symbol in the equal_weights.py code as per your requirement.
For momentum and value strategy currently mean and median can be chosen.

Disclaimer: The data is not collected from cloud and hence it is not the actual/ realtime data. The data is collected from sandbox. So it cannot be used for making any real investment decisions. This is because the token used is free token. However, paid token can be brought from https://iexcloud.io/ based on user requirement.
