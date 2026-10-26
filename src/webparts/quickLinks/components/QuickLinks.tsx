import * as React from "react";
// import styles from './QuickLinks.module.scss';
import type { IQuickLinksProps } from "./IQuickLinksProps";
// import { escape } from '@microsoft/sp-lodash-subset';
import "./style.css";
import { useEffect, useState } from "react";
import defaultIcon from "../assets/icon.png";

const QuickLinks: React.FC<IQuickLinksProps> = ({
  siteUrl,
  parentListName,
  childListName,
}) => {
  const [items, setItems] = useState<any[]>([]);
  const [links, setLinks] = useState<any[]>([]);
  const gridList = `${siteUrl}/_api/web/lists/getbytitle('${parentListName}')/items?$select=Id,Title&$expand=AttachmentFiles`;

  const linksList = `${siteUrl}/_api/web/lists/getbytitle('${childListName}')/items?`;

  useEffect(() => {
    const fetchData = async () => {
      try {
        const response1 = await fetch(gridList, {
          method: "Get",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        });
        const gridData: any = await response1.json();

        const response2 = await fetch(linksList, {
          method: "Get",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        });
        const linksData = await response2.json();

        // Process items to add the full image URL if available
        const itemsWithIcons = gridData.d.results.map((item: any) => {
          const imageUrl =
            item.AttachmentFiles.results.length > 0
              ? `${item.AttachmentFiles.results[0].ServerRelativeUrl}`
              : null;
          return { ...item, imageUrl };
        }); 
        setItems(itemsWithIcons);

        setLinks(linksData.d.results.map((link: any) => {
          return {
            childID:link.ParentId,
            text:link.URL!= null ? link.URL.Description : "",
            url:link.URL!=null? link.URL.Url : "",
          };
        }));
        console.log(linksData);

      } catch (error) {
        console.log("Error fetching data:", error); 
      }
    };
    fetchData();
  }, [siteUrl, childListName]);

  return (
    <div className="detailinks">
      {items &&
        items.map((item, index) => (
          <div className="card" key={index}>
            <div className="icon">
              <img src={item.imageUrl || defaultIcon} alt={item.Title} />
            </div>
            <h3>{item.Title}</h3> 

            {/* Filter links based on Lookups */}
            {links && links.filter((link) => link.childID === item.Id)
                .map((link, linkIndex) => (
                  <a key={linkIndex} href={link.url} className="link-item">
                    {link.text}
                  </a>
                ))}

            {links && links.filter((link) => link.childID === item.Id).length ===
                0 && <p className="no-links">No links available</p>}
          </div>
        ))}
    </div>
  );
};

export default QuickLinks;
