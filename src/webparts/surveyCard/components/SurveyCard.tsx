import * as React from "react";
import styles from "./SurveyCard.module.scss";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

const isArabic =
  window.location.pathname.toLowerCase().indexOf("/sitepages/ar/") !== -1;

interface ISurveyCardProps {
  spHttpClient: SPHttpClient;
  siteUrl: string;
  activeSurveysURL: string;
}

interface ISurveyMapped {
  id: number;
  Title: string;
  Description: string;
  SurveyURL: string;
  Created: string;
  EndDate: string;
  Remaining: string;
}

const SurveyCard: React.FC<ISurveyCardProps> = ({
  spHttpClient,
  siteUrl,
  activeSurveysURL,
}) => {
  const [survey, setSurvey] = React.useState<ISurveyMapped | null>(null);

  React.useEffect(() => {
    const fetchData = async () => {
      const response: SPHttpClientResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('Surveys')/items` +
          `?$select=Id,Title,Title_Ar,Description,Description_Ar,SurveyURL,Created,EndDate,Active` +
          `&$filter=Active eq 1` +
          `&$orderby=Created desc` +
          `&$top=1`,
        SPHttpClient.configurations.v1
      );

      const items = await response.json();

      if (items.value.length === 0) {
        setSurvey(null);
        return;
      }

      const item = items.value[0];

      // ---- Remaining Days Calculation ----
      const endDate = new Date(item.EndDate);
      const today = new Date();
      const diff = Math.ceil(
        (endDate.getTime() - today.getTime()) / (1000 * 3600 * 24)
      );

      const remaining =
        diff > 0
          ? `${diff} ${isArabic ? "يوم متبقي" : "DAYS REMAINING"}`
          : isArabic
          ? "مغلق"
          : "Closed";

      // ---- Mapping (your preferred style) ----
      const mappedItem: ISurveyMapped = {
        id: item.Id,
        Title: isArabic ? item.Title_Ar : item.Title,
        Description: isArabic ? item.Description_Ar : item.Description,
        SurveyURL: item.SurveyURL,
        Created: item.Created,
        EndDate: item.EndDate,
        Remaining: remaining,
      };

      setSurvey(mappedItem);
    };

    fetchData();
  }, [spHttpClient, siteUrl]);

  if (!survey) return <div>No active surveys</div>;

  return (
    <>
      <div
        className={isArabic ? styles.survaybgAr : styles.survaybg}
        style={{
          direction: isArabic ? "rtl" : "ltr",
          textAlign: isArabic ? "right" : "left",
        }}
      >
        <div className={styles.surveyCard}>
          <div className={styles.tag}>
            {isArabic ? "استبيان الموظفين" : "EMPLOYEE SURVEY"} -{" "}
            {survey.Remaining}
          </div>

          <h2 className={styles.title}>{survey.Title}</h2>

          <p className={styles.desc}>{survey.Description}</p>
          <div className={styles.survayActions}>
            <a
              href={survey.SurveyURL}
              style={{
                textDecoration: "none",
                color: "#fff",
                fontWeight: "bold",
              }}
              className=""
              target="_blank"
            >
              {isArabic ? " ابدأ الاستبيان ←" : "Take Survey →"}
            </a>

            <a
              href={activeSurveysURL}
              style={{
                color: "#fff",
              }}
              target="_blank"
            >
              {isArabic ? "عرض الاستبيانات النشطة" : "View Active Surveys"}
            </a>
          </div>
        </div>
      </div>
    </>
  );
};

export default SurveyCard;
