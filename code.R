library(shiny)
library(semantic.dashboard)
library(readxl)
library(dplyr)
library(ggplot2)
library(cluster)
library(caret)
library(randomForest)
library(corrplot)
library(factoextra)
library(DT)
library(rlang)
library(fmsb)
library(pls)
library(glmnet)
library(wordcloud2)
library(tm)  # 用于文本处理
library(jiebaR)
library(reshape2)

setwd("D:/download/R/item")
main_data <- read_excel("data.xlsx")
rate_data <- read_excel("individual_rate.xlsx")

# 定义 UI
ui <- dashboardPage(
  dashboardHeader(title = "国内影视作品评分预测"),
  
  dashboardSidebar(
    sidebarMenu(
      menuItem("词云图", tabName = "wordcloud_analysis"),
      menuItem("聚类分析", tabName = "cluster_analysis"),
      menuItem("相关性分析", tabName = "creator_analysis"),
      menuItem("主创分析", tabName = "director_search"),
      menuItem("模型预测", tabName = "model_prediction")
    )
  ),
  
  dashboardBody(
    tabItems(
      tabItem(tabName = "wordcloud_analysis",
              fluidRow(
                box(
                  title = "参数设置", width = 6, status = "primary", solidHeader = TRUE,style="height:400px;",
                  selectizeInput("categ_cloud", "影视作品类型", choices = c("电影" = "m", "电视剧" = "t"),multiple=TRUE),
                  selectInput("wordcloud_type", "分析目标", choices = c("影视作品名称" = "name", "题材" = "type")),
                  actionButton("generate_wordcloud", "生成词云", style = "margin-top: 15px;")
                ),
                box(
                  title = "词云图", width = 10, status = "primary", solidHeader = TRUE,
                  wordcloud2Output("wordcloud_plot")
                )
              )
      ),
      # 聚类分析页面
      tabItem(tabName = "cluster_analysis",
              fluidRow(
                box(
                  title = "参数设置", width = 6,style = "height: 400px;", status = "primary", solidHeader = TRUE,
                  selectInput("analysis_type", "分析对象", choices = c("影视作品", "主创")),
                  conditionalPanel(
                    condition = "input.analysis_type == '影视作品'",
                    selectInput("categ", "选择类别", choices = c("电影"="m", "电视剧"="t")),
                    uiOutput("type_choices")
                  ),
                  sliderInput("clusters", "选择聚类成几类",  
                              min = 2, max = 5, value = 3),
                  actionButton("cluster_btn", "执行聚类分析",style="margin-top:15px;")
                ),
                box(
                  title = "聚类可视化", width = 10,status = "primary", solidHeader = TRUE,
                  plotOutput("cluster_plot")
                ),
                box(
                  title = "聚类结果", width = 16, status = "primary", solidHeader = TRUE,
                  DTOutput("cluster_table")
                ),
                box(
                  title = "聚类中心", width = 16, status = "primary", solidHeader = TRUE,
                  DTOutput("cluster_table2")
                )
              )
      ),
      tabItem(tabName = "creator_analysis",
              fluidRow(
                box(
                  title = "参数设置", width = 6, status = "primary", solidHeader = TRUE,
                  selectInput("analysis_type", "选择分析类型", choices = c("主创指标相关系数矩阵", "影视作品评分相关性")),
                  actionButton("analyze_btn", "开始分析",style="margin-top:15px;")
                ),
                box(
                  title = "分析结果", width = 16, status = "primary", solidHeader = TRUE,
                  DTOutput("correlation_table"),
                  plotOutput("correlation_plot")
                )
              )
     
      ),
      tabItem(tabName = "director_search",
              
                box(
                  title = "参数设置", width = 6, status = "primary", solidHeader = TRUE,
                  selectInput("director_name", "选择你想分析的导演/演员", choices = rate_data$name),
                  actionButton("analyze_btn", "开始分析",style="margin-top:15px;")
                  ),
                
                box(
                  title = "作品展示", width = 16, status = "primary", solidHeader = TRUE,
                  DTOutput("作品展示_table")
                ),
                box(
                  title = "主创信息", width = 16, status = "primary", solidHeader = TRUE,
                  DTOutput("rate_data_table")
                ),
                box(
                  title = "主创偏好", width = 8, status = "primary", solidHeader = TRUE,
                  plotOutput("type_frequency_plot")
                ),
                box(
                  title = "作品趋势", width = 8, status = "primary", solidHeader = TRUE,
                  plotOutput("rate_line_plot")
                )
              ),
      # 模型预测页面
      tabItem(tabName = "model_prediction",
              fluidRow(
                box(
                  title = "模型选择", width = 6,status = "primary", solidHeader = TRUE, style="height:400px;",
                  selectInput("model_type", "选择回归模型", 
                              choices = c("基础R",  "随机森林")),
                  actionButton("train_model", "模型训练",style="margin-top:15px;"),
                  verbatimTextOutput("model_metrics"),
                  
                ),
                box(
                  title = "效果可视化", width = 10,status = "primary", solidHeader = TRUE,
                  plotOutput("model_plot")
                ),
                box(
                  title = "评分预测", width = 6, status = "primary", solidHeader = TRUE,
                  selectInput("categ", "影视类型", choices =c("电影"="m","电视剧"="t")),
                  numericInput("year", "播出年份", value = 2024, min = 1900, max = 2100),
                  selectInput("type0", "题材第一分类", choices = unique(main_data$type0)),
                  selectInput("type1", "题材第二分类", choices = unique(main_data$type1)),
                  selectInput("type2", "题材第三分类", choices = unique(main_data$type2)),
                  selectizeInput("director", "导演", choices = unique(main_data$director), options = list(create = FALSE)),
                  selectizeInput("actor1", "第一演员", choices = unique(c(main_data$actor1, main_data$actor2)), options = list(create = FALSE)),
                  selectizeInput("actor2", "第二演员", choices = unique(c(main_data$actor1, main_data$actor2)), options = list(create = FALSE)),
                  actionButton("predict_btn", "预测评分",style="margin-top:15px"),
                  
                ),
                box(
                  title = "质量评估表", width = 10, status = "primary", solidHeader = TRUE,style="height:440px;",
                  verbatimTextOutput("prediction_result"),
                  DTOutput("rating_table")
                )
                )
              )
      )
    )
  )



# 定义服务器逻辑
server <- shinyServer(function(input, output, session){
  observeEvent(input$generate_wordcloud, {
    # 根据选择的影视作品类型过滤数据
    selected_data <- main_data %>%
      filter((input$categ_cloud == "m" & grepl("m", categ)) | 
               (input$categ_cloud == "t" & grepl("t", categ)))
    
    # 根据词云类型选择数据列
    if (input$wordcloud_type == "name") {
      # 影视作品名称，需要进行分词
      word_data <- selected_data$name
      # 对名称进行分词
      jieba <- worker(bylines = TRUE)
      # 对中文和英文文本进行分词
      word_tokens <- unlist(lapply(word_data, function(x) {
        # 对中文使用jiebaR进行分词
        if (grepl("[\u4e00-\u9fa5]", x)) {
          return(segment(x, jieba))  # 对中文文本分词
        } else {
          return(strsplit(x, "\\s+")[[1]])  # 对英文文本按空格分词
        }
      }))
      # 生成词频
      word_freq <- table(word_tokens)
      word_freq_df <- as.data.frame(word_freq)
      colnames(word_freq_df) <- c("word", "freq")
      
    } else if (input$wordcloud_type == "type") {
      # 题材类型，直接计算频率
      word_data <- c(selected_data$type0, selected_data$type1, selected_data$type2)
      
      # 生成词频
      word_freq <- table(word_data)
      word_freq_df <- as.data.frame(word_freq)
      colnames(word_freq_df) <- c("word", "freq")
    }
    
    # 生成词云图
    output$wordcloud_plot <- renderWordcloud2({
      wordcloud2(word_freq_df, size = 1.5, color = "random-light", backgroundColor = "white")
    })
  })
  output$type_choices <- renderUI({
    types <- main_data %>% 
      select(type0, type1, type2) %>%
      unlist() %>% 
      unique() %>% 
      na.omit()
    selectizeInput("types", "选择类型", choices = types, multiple = TRUE)
  })
  # 更新变量选择项
  observe({
    updateSelectizeInput(session, "correlation_vars", choices = colnames(main_data), server = TRUE)
  })
  
  # 聚类分析
  observeEvent(input$cluster_btn, {
    if (input$analysis_type == "影视作品") {
      # 选择影视作品聚类变量
      data <- main_data %>% 
        filter(categ == input$categ & (type0 %in% input$types | type1 %in% input$types | type2 %in% input$types)) %>%
        select_if(is.numeric) # 选择数值型变量以进行聚类
      original_data <- main_data %>% 
        filter(categ == input$categ & (type0 %in% input$types | type1 %in% input$types | type2 %in% input$types))
    } else {
      # 选择主创聚类变量
      data <- rate_data %>% select(-name) %>% select_if(is.numeric)
      original_data <- rate_data
    }
    # 执行聚类分析
    kmeans_result<-kmeans(scale(data), centers = input$clusters)
    
    output$cluster_plot <- renderPlot({
      fviz_cluster(kmeans_result, data = data)
    })
    # 创建带聚类标签的结果表格，仅保留name和Cluster列，并按Cluster排序
    cluster_table <- original_data
    cluster_table$Cluster <- kmeans_result$cluster  # 添加聚类标签列
    # 按Cluster列排序
    cluster_table_sorted <- cluster_table %>%
      arrange(Cluster)  # 按Cluster列排序
    if (input$analysis_type != "影视作品"){
      output$cluster_table <- renderDT({
        datatable(cluster_table_sorted %>% select(name, Cluster), options = list(
          pageLength = 10,  # 每页显示5行数据
          autoWidth = TRUE,  # 自动调整列宽
          columnDefs = list(list(className = 'dt-center', targets = 0:1))  # 中心对齐name和Cluster列
        ))
      })
    }else{
      output$cluster_table <- renderDT({
        datatable(cluster_table_sorted %>% select(name, rate, Cluster), options = list(
          pageLength = 10,  # 每页显示5行数据
          autoWidth = TRUE,  # 自动调整列宽
          columnDefs = list(list(className = 'dt-center', targets = 0:1))  # 中心对齐name和Cluster列
        ))
      })
    }
    cluster_centers <- kmeans_result$centers
    # 生成聚类结果的数据框
    cluster_results <- data.frame(
      Name = rownames(cluster_centers),
      Center = cluster_centers
    )
    # 输出聚类结果表格
    output$cluster_table2 <- renderDT({
      datatable(cluster_results, options = list(scrollX=TRUE,pageLength = 5))
    })
    print(cluster_results)
  })
  
  # 相关性分析
  # 当点击开始分析按钮时执行分析
  observeEvent(input$analyze_btn, {
    
    if (input$analysis_type == "主创指标相关系数矩阵") {
      # 设置分析的变量
      correlation_data <- rate_data[, c("director_m", "director_t", "actor_1m", "actor_2m",
                                        "actor_1t", "actor_2t", "m_p", "m_tgi", "w_p", "w_tgi")]
      
      # 计算相关系数矩阵
      cor_matrix <- cor(correlation_data, use = "complete.obs")
      
      # 输出相关系数矩阵
      output$correlation_table <- renderDT({
        datatable(as.data.frame(cor_matrix),options = list(scrollX=TRUE,pageLength = 10))
      })
      
      # 绘制相关性热图
      output$correlation_plot <- renderPlot({
        cor_matrix <- cor(correlation_data, use = "complete.obs")
        cor_matrix_melted <- melt(cor_matrix, varnames = c("Var1", "Var2"), value.name = "Correlation")
        
        # 绘制相关性矩阵热图
        ggplot(cor_matrix_melted, aes(Var1, Var2, fill = Correlation)) +
          geom_tile() +  # 绘制热图
          scale_fill_gradient2(low = "blue", high = "red", mid = "white", midpoint = 0, limit = c(-1, 1), name="Correlation") +
          theme_minimal() + 
          theme(axis.text.x = element_text(angle = 45, hjust = 1), # x轴标签倾斜
                axis.text.y = element_text(angle = 0)) +
          labs(x = "Variables", y = "Variables", title = "Correlation Matrix")
      })
      
      
    } else if (input$analysis_type == "影视作品评分相关性") {
      # 设置分析的变量
      correlation_data <- main_data[, c("rate", "year", "director_m", "director_t", 
                                        "d_m_p", "d_m_tgi", "d_w_p", "d_w_tgi", 
                                        "actor1_m", "actor1_t", "a1_m_p", "a1_m_tgi", 
                                        "a1_w_p", "a1_w_tgi", "actor2_m", "actor2_t", 
                                        "a2_m_p", "a2_m_tgi", "a2_w_p", "a2_w_tgi")]
      
      # 计算相关系数矩阵
      cor_matrix <- cor(correlation_data, use = "complete.obs")
      
      # 输出相关系数矩阵
      output$correlation_table <- renderDT({
        datatable(as.data.frame(cor_matrix),options = list(scrollX=TRUE,pageLength = 20))
      })
      
      # 绘制相关性热图
      output$correlation_plot <- renderPlot({
        cor_matrix <- cor(correlation_data, use = "complete.obs")
        cor_matrix_melted <- melt(cor_matrix, varnames = c("Var1", "Var2"), value.name = "Correlation")
        
        # 绘制相关性矩阵热图
        ggplot(cor_matrix_melted, aes(Var1, Var2, fill = Correlation)) +
          geom_tile() +  # 绘制热图
          scale_fill_gradient2(low = "blue", high = "red", mid = "white", midpoint = 0, limit = c(-1, 1), name="Correlation") +
          theme_minimal() + 
          theme(axis.text.x = element_text(angle = 45, hjust = 1), # x轴标签倾斜
                axis.text.y = element_text(angle = 0)) +
          labs(x = "Variables", y = "Variables", title = "Correlation Matrix")
      })
    }
  })
  observeEvent(input$analyze_btn, {
    selected_director <- input$director_name
    # 筛选 main_data 中导演、演员包含该主创的所有作品
    filtered_main_data <- main_data %>%
      filter(director == selected_director | actor1 == selected_director | actor2 == selected_director)
    # 显示“作品展示”表格
    output$作品展示_table <- renderDT({
      datatable(filtered_main_data %>% select(name, director, actor1, actor2, rate),
                options = list(pageLength = 10, autoWidth = TRUE))
    })
    
    # 显示对应的 rate_data 中的所有列
    selected_rate_data <- rate_data %>%
      filter(name == selected_director) %>%
      mutate(
        director_m = round(director_m, 2),
        director_t = round(director_t, 2),
        actor_1m = round(actor_1m, 2),
        actor_2m = round(actor_2m, 2),
        actor_1t = round(actor_1t, 2),
        actor_2t = round(actor_2t, 2)
      )
    output$rate_data_table <- renderDT({
      datatable(selected_rate_data, options = list(pageLength = 1,scrollX = TRUE, columnDefs = list(list(className = 'dt-center', targets = "_all"))))
    })
    print(selected_rate_data)
    library(tidyr)
    # 绘制雷达图：统计筛选后的 main_data 中 type、type1、type2 的值的频率
    type_counts <- filtered_main_data %>%
      select(type0, type1, type2) %>%
      pivot_longer(cols = everything(), names_to = "type_column", values_to = "type_value") %>%
      filter(!is.na(type_value)) %>%
      count(type_value)

    # Step 2: 雷达图绘制
    output$type_frequency_plot <- renderPlot({
      # 使用 pivot_wider() 代替 spread()
      radar_data <- type_counts %>%
        pivot_wider(names_from = type_value, values_from = n, values_fill = list(n = 0))
        # 删除不需要的列
      
      # 确保数据框格式适合雷达图
      radar_data <- as.data.frame(radar_data)
      radar_data <- rbind(rep(10, ncol(radar_data)), rep(0, ncol(radar_data)), radar_data)
      
      # 绘制雷达图
      radarchart(radar_data)
      
      output$rate_line_plot <- renderPlot({
        # 绘制折线图，x 为 year，y 为 rate
        ggplot(filtered_main_data, aes(x = year, y = rate)) +
          geom_line(color = "blue") +  # 使用蓝色折线
          geom_point(color = "red") +   # 使用红色点标记数据点
          labs( x = "年份", y = "评分") +  # 添加标题和轴标签
          theme_minimal() +  # 使用简洁的主题
          theme(axis.text.x = element_text(angle = 45, hjust = 1))  # 使x轴标签倾斜
      })
    })
  })
  # 训练模型
  observeEvent(input$train_model, {
    # train_data <- read_xlsx("train.xlsx")
    # train_data <- train_data %>% select(-name,-director,-actor1,-actor2)
    # test_data <- read_xlsx("test.xlsx")
    # test_data <- test_data %>% select(-name,-director,-actor1,-actor2)
    set.seed(123)
    # 创建训练集和测试集的索引
    train_index <- createDataPartition(main_data$rate, p = 0.95, list = FALSE)
    # 分割数据为训练集和测试集
    train_data <- main_data[train_index, ]
    test_data <- main_data[-train_index, ]
    # 移除不需要的列（name, director, actor1, actor2）
    train_data <- train_data %>% select(-name, -director, -actor1, -actor2)
    test_data <- test_data %>% select(-name, -director, -actor1, -actor2)
    # 根据用户选择的回归模型类型，选择不同的训练方法
    if(input$model_type == "基础R"){
      model<-lm(rate ~ ., data = train_data)
      predictions<- predict(model, newdata = test_data)
      mse <- mean((test_data$rate - predictions)^2)
      # 计算均方根误差（RMSE）
      rmse <- sqrt(mse)
      # 计算R平方值（R²）
      r_squared <- summary(model)$r.squared
      model<-lm(rate ~ ., data =  main_data%>% select(-name))
    }else{
        model<- randomForest(rate ~ ., data = train_data, mtry = 10, ntree = 100)
        predictions <- predict(model, newdata = test_data)
        mse <- mean((test_data$rate - predictions)^2)
        rmse <- sqrt(mse)
        r_squared <- cor(test_data$rate, predictions)^2
        model<- randomForest(rate ~ ., data = main_data%>% select(-name), mtry = 10, ntree = 100)
      }
    # 输出模型性能指标：MSE、RMSE 和 R²
    output$model_metrics <- renderPrint({
        cat("MSE:", mse, "\n")
        cat("RMSE:", rmse, "\n")
        cat("R²:", r_squared, "\n")
    })
    output$model_plot <- renderPlot({
      # 假设 train_data 是训练集，rate 是目标变量
      plot_data <- data.frame(Actual = test_data$rate, Predicted = predictions)
      ggplot(plot_data, aes(x = Actual, y = Predicted)) +
      geom_point(color = "blue", size = 3, alpha = 0.6) +  # 散点图，点的大小和透明度
      geom_abline(slope = 1, intercept = 0, color = "red", linetype = "dashed") +  # 添加完美预测线
      geom_smooth(method = "lm", se = FALSE, color = "green", linetype = "solid") +  # 添加回归线
      labs(x = "Actual Rating", y = "Predicted Rating", title = "Comparison") +
      theme_minimal() +  # 使用简洁主题
      theme(
        plot.title = element_text(hjust = 0.5, size = 16),  # 标题居中，大小调整
        axis.title = element_text(size = 14),  # 坐标轴标题的字体大小
        axis.text = element_text(size = 12),  # 坐标轴刻度的字体大小
        panel.grid.major = element_line(color = "gray", size = 0.5),  # 主网格线样式
        panel.grid.minor = element_line(color = "gray", size = 0.25, linetype = "dotted")  # 次网格线样式
      )
      # 添
    })
  })
  rating_data <- data.frame(
    range = c("8.5-10", "7.5-8.5", "6.5-7.5", "6-6.5", "2-6"),
    level = c("H", "MH", "M", "ML", "L"),
    advice = c("极佳", "好", "中", "差", "极差"),
    description = c("极具投资价值", "值得投资", "需要认真考虑", "有投资风险", "高投资风险"),
    stringsAsFactors = FALSE
  )
    observeEvent(input$predict_btn, {
      # 从选中的数据中提取特征
      director_data <- rate_data[rate_data$name == input$director, ]
      actor1_data <- rate_data[rate_data$name == input$actor1, ]
      actor2_data <- rate_data[rate_data$name == input$actor2, ]
      # 提取对应的特征
      input_features <- list(
        categ = input$categ,
        year = input$year,
        type0 = input$type0,
        type1 = input$type1,
        type2 = input$type2,
        # director = input$director,
        director_m = director_data$director_m[1],  # 只取第一个匹配的导演
        director_t = director_data$director_t[1],
        d_age_p = director_data$age_p[1],
        d_age_tgi = director_data$age_tgi[1],
        d_m_p = director_data$m_p[1],
        d_m_tgi = director_data$m_tgi[1],
        d_w_p = director_data$w_p[1],
        d_w_tgi = director_data$w_tgi[1],
        # actor1 = input$actor1,
        actor1_m = actor1_data$actor_1m[1],
        actor1_t = actor1_data$actor_1t[1],
        a1_age_p = actor1_data$age_p[1],
        a1_age_tgi = actor1_data$age_tgi[1],
        a1_m_p = actor1_data$m_p[1],
        a1_m_tgi = actor1_data$m_tgi[1],
        a1_w_p = actor1_data$w_p[1],
        a1_w_tgi = actor1_data$w_tgi[1],
        # actor2 = input$actor2,
        actor2_m = actor2_data$actor_2m[1],
        actor2_t = actor2_data$actor_2t[1],
        a2_age_p =actor2_data$age_p[1],
        a2_age_tgi = actor2_data$age_tgi[1],
        a2_m_p = actor2_data$m_p[1],
        a2_m_tgi = actor2_data$m_tgi[1],
        a2_w_p = actor2_data$w_p[1],
        a2_w_tgi = actor2_data$w_tgi[1]
      )
      
      input_features <- data.frame(input_features)
      if(input$model_type == "基础R"){
        model<-lm(rate ~ ., data =  main_data%>% select(-name,-director,-actor1,-actor2))
      }else{
        model<- randomForest(rate ~ ., data = main_data%>% select(-name,-director,-actor1,-actor2), mtry = 10, ntree = 100)
      }
      # 使用模型进行预测
      prediction <- predict(model, newdata = input_features)
      
      # 输出预测结果
      rating_result <- rating_data[apply(rating_data, 1, function(row) {
        pred_range <- as.numeric(strsplit(row["range"], "-")[[1]])
        prediction >= pred_range[1] & prediction <= pred_range[2]
      }), ]
      
      # 输出预测结果
      output$prediction_result <- renderText({
        paste("Prediction: ", prediction, 
              "\nLevel: ", rating_result$level, 
              "\nAdvice: ", rating_result$advice, 
              "\nDescription: ", rating_result$description)
      })
      # 渲染表格
      
      
})
    
    output$rating_table <- renderDT({
      datatable(rating_data, options = list(pageLength = 5, autoWidth = TRUE))
    })
    
})

# 运行应用
shinyApp(ui = ui, server = server)
