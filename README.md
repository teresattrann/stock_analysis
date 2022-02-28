# VBA Challenge

## Overview of Project

### Background

Steve would like to expand his dataset to include the entire stock market for the last few years. Currently, Steve is only working with a dozen amount of stocks. He would like to now expand his dataset to include thousands of stocks and we would like to help him.

### Purpose

The purpose of this challenge is to refactor our existing code in order to loop through all our data once and collect information. Currently, our code is only able to work for a small amount of stocks, and not large amounts. By refactoring, we are able to receive the same output that can hold more  data, while using less and more readable code, therefore, making it more efficient. 

## Results

### Analysis

According to our data, the stock performance for 2018 seem to drop drastically from 2017. As you can see, every stock in 2017, with the exception of TERP, had a very high return value, whereas in 2018, the same return value was in the negatives. 

![VBA_Challenge2017](https://user-images.githubusercontent.com/98780937/156061903-c11a5f3b-84cf-41a9-90a2-281c56242b6d.png)
![VBA_Challenge2018](https://user-images.githubusercontent.com/98780937/156061929-1a4e2b57-3322-495a-8795-4dba048c6675.png)

The execution time however, seemeed to improve exceptionally. Given our data on execution time, the refactored code seem to work 4 to 5 times facter than the original non refactored code. 

Refactored code execution time:


![2017_RunTime](https://user-images.githubusercontent.com/98780937/156062873-26ef80e4-27fa-49ae-a26d-5bd49a01cbea.png)
![2018_RunTime](https://user-images.githubusercontent.com/98780937/156062888-07cc7ab3-9570-49a0-af5d-18de24449464.png)

Original code execution time:


![Original2017](https://user-images.githubusercontent.com/98780937/156063073-06ab4722-d165-414b-a93a-be64f94de57f.png)
![Original2018](https://user-images.githubusercontent.com/98780937/156063082-c7abe485-1a2c-42cf-86ab-1dc62fba0084.png)

## Summary

### Advantages and Disadvantages of Refactoring Code

The advantages of refactoring code is to basically improving and make the code more efficient. By efficient, that can mean using less code, as well as making the code more readable for future users. By using less code, we can also cut down on memory as well. The disadvantages of refactoring code is that we could potentially mess up our already perfectly working code. 

### Advantages and Disadvantages of Original and Refactored VBA Script

Our refactored VBA Script seemed to have more advantages than disadvantages. Our biggest advantage was the execution time. We had about a 4 to 5 time increase in efficiency (macro run time) when we refactored our code. We were also able to clean up and and make our code a lot smoother.
