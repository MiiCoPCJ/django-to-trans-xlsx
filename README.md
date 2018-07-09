安装requirement需要的库

测试office环境：office 2007 sp2

可能需要相关插件
SaveAsPDFandXPS.exe
http://download.microsoft.com/download/6/2/5/6259b99f-1abf-4f27-b2a0-ad018b04f0a6/SaveAsPDFandXPS.exe


要转换的django的model文件放在resource文件夹
model默认格式

class CoinDetail(models.Model):
    RECHARGE = 1
    REALISATION = 2
    TYPE_CHOICE = (
        (RECHARGE, "充值"),
        (REALISATION, "提现"),
    )
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    coin_name = models.CharField(verbose_name="货币名称", max_length=255, default='')
    amount = models.CharField(verbose_name="操作数额", max_length=255)
    rest = models.DecimalField(verbose_name="余额", max_digits=10, decimal_places=3, default=0.000)
    sources = models.CharField(verbose_name="资金流动类型", choices=TYPE_CHOICE, max_length=1, default=BETS)
    is_delete = models.BooleanField(verbose_name="是否删除", default=False)
    created_at = models.DateTimeField(verbose_name="操作时间", auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = verbose_name_plural = "用户资金明细"

执行todo文件夹trans_v2
自动生成 .docx   /doc/word
        .pdf    /doc/pdf