# douban_spider
爬取豆瓣下任何标签的书籍标题，评分，评论人数等信息
### DouBan Crawler Series 
> 完成豆瓣读书/电影相关的爬取，使用了简单的**多线程**极大地提高了爬虫效率，更多信息待加入补充。

### 豆瓣图书爬虫    [Python 3.6.1]
> 爬取结果在`Result_Book`文件夹，可直接查看  <br>

#### 实现功能： 
 - 增加了简单的多线程爬取(concurrent.futures模块，线程池管理)，极大地提高了爬虫效率。
 - 按标签名称进行相关图书信息的抓取，排序后存入本地excel，可自行进行进一步筛选，按`Tag`存取在不同的`Sheet`
 - 使用`User Agent`伪装成不同的浏览器进行爬取，并加入随机延时来更好的模仿浏览器行为，避免爬虫被封
 - 在cimonCqk代码的基础上，新增了“评论人数”的标签，但是新增了这个标签后会造成douban的IP封杀，因为'评论人数'需要用到递归，想在接下来用代理池的方法解决这个问题，待成功后再更新。
    
##### 豆瓣页面截图：

![Page](https://github.com/SimonCqk/DouBanCrawls/blob/master/ScreenShots/reading_page.jpg?raw=true)

##### 运行时截图：

![Running](https://github.com/SimonCqk/DouBanCrawls/blob/master/ScreenShots/reading_running.jpg?raw=true)

##### Excel结果截图：

![Excel](https://github.com/SimonCqk/DouBanCrawls/blob/master/ScreenShots/reading_excel.jpg?raw=true)

### 豆瓣电影爬虫
> 爬取结果在`Result_Movie`文件夹，可直接查看 <br>
#### 实现功能： 
 - 增加了简单的多线程爬取(threading模块，简单粗暴的多线程管理)，极大地提高了爬虫效率。
 - 按标签名称进行相关电影信息的抓取，排序后存入本地excel，可自行进行进一步筛选，按`Tag`存取在不同的`Sheet`
 - 使用`User Agent`伪装成不同的浏览器进行爬取，并加入随机延时来更好的模仿浏览器行为，避免爬虫被封
 
   
##### 豆瓣页面截图：

![Page](https://github.com/SimonCqk/DouBanCrawls/blob/master/ScreenShots/movie_page.jpg?raw=true)

##### 运行时截图：

![Running](https://github.com/SimonCqk/DouBanCrawls/blob/master/ScreenShots/movie_running.jpg?raw=true)

##### Excel结果截图：

![Excel](https://github.com/SimonCqk/DouBanCrawls/blob/master/ScreenShots/movie_excel.jpg?raw=true)


> 欢迎 Star / PR.
