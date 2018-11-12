do
	a = msgbox(" 我 其 实 真 的 很 喜 欢 你 !", vbYesNoCancel, "你 愿 意 做 我 的 女 朋 友 吗 ？")
	if a = vbYes then 
		msgbox " 以 后 我 们 将 共 度 每 一 天 ，分 享 每 一 天 的 快 乐 、悲伤",  ," 恭 喜 你 成 为 我 的 女 朋 友 " 
		exit do
	elseif a = vbNo then
		b = msgbox (" 有 点 点 伤 心 ， 有 点 点 难 过 ， 不 要 再 点 是 了 !",  vbYesNo, " 你 不 愿 意 吗 ？ "  )
		if b = vbYes then 
			msgbox " 可 是 很 气 的 是 ， 你 点 “是” 也 关 闭 不 了 啊 ！" ,  , " 难 道 你 就 这 样 讨 厌 我 吗 ？" 
		else
			msgbox " 还 是 有 那 么 一 点 点 可 爱 不 是 嘛 ！",," 我 就 知 道 你 还 是 舍 不 得 我 的 ！"
		end if
	else
		d= msgbox(" 既 不 喜 欢 我 ， 也 不 讨 厌 我 吗 ？ " , vbYesNo ," 取 消 是 什 么 意 思 呢 ？ " )
		if d = vbYes then 
			msgbox " 给 我 一 个 机 会 吧 ！",,"可 是 我 是 很 喜 欢 你 呢 ！" 

		else
			msgbox " 还 是 乖 乖 的 接 收 吧 ！",," 嘿嘿 ， 不 要 想 关 掉 我 哦 ！ "
		end if
	end if
loop