Function Geolocation_Get_Distance(ByVal startLongitude As Double, ByVal startLatitude As Double, ByVal endLongitude As Double, ByVal endLatitude As Double) as Double
        Dim return_value As Double = 0
        Try
            Dim d2r = 3.14159265359 / 180          ' Множитель для перевода градусов в радианы
            Dim major = 6378137.0           ' Большая полуось
            Dim minor = 6356752.314245   ' Малая полуось
            Dim e2 = 0.006739496742337  ' Площадь эксцентриситета эллипсоида  
            Dim flat = 0.003352810664747 ' Свед`ение эллипсоида
            ' Получаем разницы между широтами-долготами
            Dim lambda = (startLongitude - endLongitude) * d2r          ' Разность долгот
            Dim phi = (startLatitude - endLatitude) * d2r                     ' Разность широт
            Dim phiMean = ((startLatitude + endLatitude) / 2.0) * d2r  'Средняя широта
            ' Расчет мередианального и траверсного радиусов кривизны в средних широтах  
            Dim temp = 1 - e2 * (Math.Pow(Math.Sin(phiMean), 2.0)) 'Временная переменная
            Dim rho = (major * (1 - e2)) / Math.Pow(temp, 1.5)            ' Меридиональный радиус кривизны
            Dim nu = major / Math.Sqrt(1 - e2 * Math.Pow(Math.Sin(phiMean), 2.0))  ' Поперечный РК
            ' Расчет углового расстояния
            Dim z = Math.Sqrt(Math.Pow(Math.Sin(phi / 2.0), 2.0) + Math.Cos(endLatitude * d2r) * Math.Cos(startLatitude * d2r) * Math.Pow(Math.Sin(lambda / 2.0), 2.0))
            z = 2 * Math.Asin(z)         ' Угловое расстояние в центре сфероида
            ' Расчет азимута
            Dim alpha = Math.Cos(endLatitude * d2r) * Math.Sin(lambda) * 1 / Math.Sin(z)
            alpha = Math.Asin(alpha)  ' Азимут
            ' Используем Теорему Эйлера для определения радиуса сферической Земли
            Dim r = rho * nu / (rho * Math.Pow(Math.Sin(alpha), 2.0) + nu * Math.Pow(Math.Cos(alpha), 2.0))
            ' Устанавливаем азимут и расстояние
            ret = z * r ' Дистанция
        Catch ex As Exception

        End Try
        Return return_value
End Function